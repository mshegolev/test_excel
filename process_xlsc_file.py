import zipfile
from lxml import etree
import os
import copy


def uzip_xlsx_document(template_path, unzip_path):
    zip_ref = zipfile.ZipFile(template_path, 'r')
    zip_ref.extractall(unzip_path)
    zip_ref.close()
    return unzip_path


def get_xml_tree_with_file_content(unzipped_path):
    xml_filepath = get_xml_path(unzipped_path)
    with open(xml_filepath, 'rb') as file_handler:
            xml_str = file_handler.read()
    return etree.fromstring(xml_str)


def process_xml_tree_with_context(base_xml_tree, context):
    xml_tree = copy.deepcopy(base_xml_tree)
    for text_node in get_all_text_nodes(xml_tree):
        for var_name, var_val in context.items():
            var_placeholder_name = '{{ %s }}' % var_name
            if var_placeholder_name in text_node.text:
                text_node.text.replace(
                    var_placeholder_name,
                    str(var_val),
                )
    return xml_tree


def get_all_text_nodes(xml_tree):
    return xml_tree.findall(
        '..//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')


def save_xml_tree_to_file(processed_xml_tree, unzipped_path):
    xml_filepath = get_xml_path(unzipped_path)
    with open(xml_filepath, 'wb') as file_handler:
        file_handler.write(
            etree.tostring(
                processed_xml_tree
                )
            )


def get_xml_path(unzipped_path):
    return os.path.join(unzipped_path, 'xl', 'sharedStrings.xml')


def zip_document_to_xlsx(unzipped_path, result_file_path):
    zip_ref = zipfile.ZipFile(result_file_path, 'w', zipfile.ZIP_DEFLATED)
    for root, dirs, files in os.walk(unzipped_path):
        for file in files:
            zip_ref.write(
                os.path.join(root, file),
                os.path.relpath(os.path.join(root, file), unzipped_path)
                )
    zip_ref.close()


def process_xlsx_template(template_path, result_file_path,
                          context, unzip_path='/tmp/excel_tmp'):
    unzipped_path = uzip_xlsx_document(template_path, unzip_path)
    xml_tree = get_xml_tree_with_file_content(unzipped_path)
    processed_xml_tree = process_xml_tree_with_context(xml_tree, context)
    save_xml_tree_to_file(processed_xml_tree, unzipped_path)
    zip_document_to_xlsx(unzipped_path, result_file_path)


if __name__ == '__main__':
    process_xlsx_template(
        template_path='/Users/user/Documents/develop' +
        '/excel.xlsx',
        result_file_path='/Users/user/Documents/develop' +
        '/excel_result.xlsx',
        context={
            'student_full_name': 'My name',
            'course_name': 'web developer go',
            'course_finish_date': '01.01.18',
            'challenges_done': 20,
            'challenges_total': 20,
            }
    )
