import zipfile
import shutil
import re
import os


def unzip_folder(zip_file_path, extract_folder_path):
    """Unzip the xlsx file to the target folder"""
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_folder_path)

def handle_file(extracted_folder, vba_path=None):
    """Process unzipped files"""
    if vba_path:
        # 1、Copy the macro binary file to the target folder
        shutil.copy(vba_path, extracted_folder+"/xl")
        # 2、modify xl/_rels/workbook.xml.rels
        modify_workbook_xml_rels(extracted_folder+"/xl/_rels/workbook.xml.rels")
    # 3、modify [Content_Types].xml
    modify_content_xml(extracted_folder+"/[Content_Types].xml")
    # 4、modify xl/workbook.xml
    modify_workbook_xml(extracted_folder+"/xl/workbook.xml")

def modify_content_xml(xml_file):
    with open(xml_file, 'r', encoding="utf-8") as file:
        xml_content = file.read()

    # Use regular expressions to match <Default Extension content
    pattern = r"<Default Extension(.*?)>"
    insert_text = '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
    modified_string = re.sub(pattern, lambda match: match.group() + insert_text, xml_content, flags=re.DOTALL, count=1)
    # Use regular expressions to match <Override PartName="/xl/workbook.xml" content
    pattern = r'<Override PartName="/xl/workbook.xml"(.*?)>'
    sub_text = '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>'
    modified_string = re.sub(pattern, sub_text, modified_string, flags=re.DOTALL, count=1)
    # Write modified content to new XML file
    with open(xml_file, 'w', encoding="utf-8") as file:
        file.write(modified_string)

def modify_workbook_xml(xml_file):
    with open(xml_file, 'r', encoding="utf-8") as file:
        xml_content = file.read()

    # Use regular expressions to match <workbookPr/> content
    pattern = r"<workbookPr(.*?)>"
    sub_text = '<workbookPr codeName="ThisWorkbook"/>'
    modified_string = re.sub(pattern, sub_text, xml_content, flags=re.DOTALL, count=1)
    # Write modified content to new XML file
    with open(xml_file, 'w', encoding="utf-8") as file:
        file.write(modified_string)

def modify_workbook_xml_rels(xml_file):
    with open(xml_file, 'r') as file:
        xml_content = file.read()

    # Find all numbers starting with "rId" using regular expression
    pattern = r'rId(\d+)'
    matches = re.findall(pattern, xml_content)
    # Convert the matched numbers to integers and find the maximum value
    max_number = max(map(int, matches))
    result = str(max_number + 1)
    # Use regular expressions to match <Relationship content
    pattern = r"<Relationship(.*?)>"
    insert_text = f'<Relationship Id="rId{result}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
    modified_string = re.sub(pattern, lambda match: match.group() + insert_text, xml_content, flags=re.DOTALL, count=1)
    # Write modified content to new XML file
    with open(xml_file, 'w') as file:
        file.write(modified_string)

def zip_folder(extract_folder_path, zip_file_name):
    """Compress files whose target path is xlsm suffix"""
    with zipfile.ZipFile(zip_file_name, 'w') as zip_file:
        for root, dirs, files in os.walk(extract_folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                zip_file.write(file_path, os.path.relpath(file_path, extract_folder_path))
    print(f'The folder was successfully compressed to {zip_file_name} !')

def xlsx2xlsm(xlsx_file_path, vba_path=None):
    """Input in the xlsx file path, optional vbaProject.bin path"""
    extract_folder_path = 'extracted_folder'
    unzip_folder(xlsx_file_path, extract_folder_path)
    handle_file(extract_folder_path, vba_path)
    zip_file_name = xlsx_file_path.split(".xlsx")[0]
    zip_folder(extract_folder_path, zip_file_name+".xlsm")


if __name__ == "__main__":
    # vba_path = "vbaProject.bin"
    xlsx_file_path = '1.xlsx'
    xlsx2xlsm(xlsx_file_path)