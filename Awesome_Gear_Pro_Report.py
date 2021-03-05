from glob import iglob
from pandas import read_table, read_csv, DataFrame, ExcelWriter
from os import path, startfile
from dataclasses import make_dataclass




# Define Directory to work with
# Read the directory information from the Config File
Config_File = open('Awesome_Gear_Pro_Report_Config.txt', 'r')
Config_File = open('Awesome_Gear_Pro_Report_Config.txt', 'r')
Lines = Config_File.read().splitlines()
DIRNAME = Lines[0].split("=")[1]
INPUT_FILE = Lines[1].split("=")[1]
TEMPLATE_DIR = Lines[2].split("=")[1]
mystr = DIRNAME + INPUT_FILE

# Generate the file list
# Then Load a Source Data File


myfileList = []
for file_name in iglob(mystr, recursive=True):
    myfileList.append(file_name)
    # print(file_name)
# Get the latest modified file
latest_file = max(myfileList, key=path.getmtime)
prt = read_table(latest_file, sep='\t')
SOURCE = prt
PARTNB = SOURCE.get("partnb")[1]
PROGRAMNAME = SOURCE.get("planid")[1]


# Load the template file
TEMPLATE = read_csv(TEMPLATE_DIR)
NUM_OF_ROW = TEMPLATE.shape[0]


def cal_value(func_name, arg_list):
    """
    cal_value reads the actual value from the SOURCE text file at each of the line numbers
    given by arg_list, calculate and return the calculated result based on the func_name;
    """
    result = 0
    arg_val = [SOURCE.at[i, 'actual'] for i in arg_list]
    if func_name == "AVERAGE":
        result = sum(arg_val) / len(arg_val)
    if func_name == "EQUAL":
        result = sum(arg_val)
    if func_name == "MIN":
        result = min(arg_val)
    if func_name == "MAX":
        result = max(arg_val)

    return round(result, 4)


def cal_status(actual_value, usl=100, lsl=-100):
    """
    cal_status return "NOT OK" if actual_value is out side of usl and lsl;
    """
    status = ""
    if actual_value > usl or actual_value < lsl:
        status = "NOT OK"
    return status


# print(TEMPLATE)

# Define Dimension class to represent a dimension
Dimension = make_dataclass("Dimension",
                           [("Label", str), ("Description", str), ("USL", str), ("LSL", str), ("Actual", float),
                            ("Status", str)])
dim_list = []
out_of_tolerance_list=[]


for r in range(NUM_OF_ROW):
    argument = TEMPLATE.loc[r, 'Argument']
    upper_limit = TEMPLATE.loc[r, 'USL']
    lower_limit = TEMPLATE.loc[r, 'LSL']
    function_name = TEMPLATE.loc[r, 'Function']
    argument_list = list(map(int, argument.split("-")))
    value = cal_value(function_name, argument_list)
    status = cal_status(value, upper_limit, lower_limit)
    # build a dimension class
    Label = TEMPLATE.loc[r, 'Label']
    Description = TEMPLATE.loc[r, 'Description']
    USL = TEMPLATE.loc[r, 'USL']
    LSL = TEMPLATE.loc[r, 'LSL']
    Actual = value
    Status = status
    d = Dimension(Label, Description, USL, LSL, Actual, Status)
    dim_list.append(d)

    if Status=="NOT OK":
        out_of_tolerance_list.append(r)
    # print(TEMPLATE.loc[r, 'Function'], TEMPLATE.loc[r, 'Argument'], value, status)

result = DataFrame(dim_list)
# print(result)


# write the result to a excel file
saved_file_name = path.basename(latest_file)
saved_file_name = DIRNAME+saved_file_name

writer = ExcelWriter(saved_file_name+".xlsx")
result.to_excel(writer, 'Sheet1')
workbook1 = writer.book
worksheets = writer.sheets
worksheet1 = worksheets['Sheet1']
# Make the description column longer
worksheet1.set_column("C:C", 35)
# worksheet1.column_dimensions['C'].width = 35
format1 = workbook1.add_format({'bold':  True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
format2 = workbook1.add_format({'bold':  True, 'align': 'left', 'valign': 'top', 'text_wrap': False})


for line_num in out_of_tolerance_list:
    worksheet1.set_row(line_num+1, cell_format=format1)


text1 = '             '
text2 = 'Tolerance File is at ' + TEMPLATE_DIR

text3 = 'Source Data File is at ' + saved_file_name
worksheet1.write(len(result) + 1, 0, text1)
worksheet1.write(len(result) + 2, 0, text2)
# worksheet1.write(len(result) + 3, 0, text3)

text4 = PROGRAMNAME + "     " + str(PARTNB)

worksheet1.write(len(result) + 4, 0, text4)
worksheet1.set_row(len(result) + 4, cell_format=format2)

text5 = 'This part is OK'
if len(out_of_tolerance_list) > 0:
    text5 = 'This part is NOT OK'
worksheet1.write(len(result) + 5, 0, text5)
worksheet1.set_row(len(result) + 5, cell_format=format2)
writer.save()

# Print the result from the absolute path
startfile(path.abspath(saved_file_name+".xlsx"), "print")