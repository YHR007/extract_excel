import os
import traceback
import shutil
import zipfile
import tkinter as tk
from tkinter import filedialog, ttk
from xml.etree import ElementTree as ET
from MRtoDocx import MRtoDocx


class StructuredExcelProcessor:
    def __init__(self):
        self.namespaces = {
            'nsr':'http://schemas.openxmlformats.org/package/2006/relationships',
            'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

    def process_excel(self, excel_path, output_dir, image_dir):
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        os.makedirs(image_dir, exist_ok=True)

        # 处理图片和文本数据
        workbook_data = self.process_workbook(excel_path, image_dir)

        # 保存知识库
        MRtoDocx(workbook_data, output_dir)




    def process_workbook(self, excel_path, image_dir):
        temp_dir = 'temp_excel'
        with zipfile.ZipFile(excel_path, 'r') as z:
            z.extractall(temp_dir)

        data_mapping = {}
        workbook_rels = self.process_workbook_rels(temp_dir)
        for root, _, files in os.walk(os.path.join(temp_dir, 'xl/worksheets')):
            for file in files:
                if file.startswith('sheet') and file.endswith('.xml'):
                    sheet_name = self.get_sheet_name(file, temp_dir, workbook_rels)
                    image_mapping=self.process_sheet_images(root, file, temp_dir, image_dir)
                    text_mapping=self.process_sheet_texts(root, file, temp_dir)
                    data_mapping[sheet_name]=image_mapping
                    data_mapping[sheet_name].update(text_mapping)

        shutil.rmtree(temp_dir)
        return data_mapping

    def process_workbook_rels(self, temp_dir):
        workbook_rel_tree = ET.parse(os.path.join(temp_dir, 'xl/_rels/workbook.xml.rels'))
        return {rel.get('Target').split('/')[-1]: rel.get('Id')
                for rel in workbook_rel_tree.findall('.//nsr:Relationship', self.namespaces)}

    def get_sheet_name(self, xml_file, temp_dir,workbook_rels):
        workbook_tree = ET.parse(os.path.join(temp_dir, 'xl/workbook.xml'))
        sheets = workbook_tree.findall('.//ns:sheets/ns:sheet', self.namespaces)
        _id=workbook_rels[xml_file]

        for s in sheets:
            rid=s.get('{%s}id' % self.namespaces['r'])
            if rid==_id:
                return s.get('name')

    def process_sheet_images(self, root, xml_file, temp_dir, image_dir):
        tree = ET.parse(os.path.join(root, xml_file))
        drawing_ref = tree.find('.//ns:drawing', self.namespaces)
        image_mapping={}

        if drawing_ref is not None:
            drawing_id = drawing_ref.get('{%s}id' % self.namespaces['r'])
            rels_path = os.path.join(temp_dir, 'xl/worksheets/_rels', f'{xml_file}.rels')
            drawing_file = self.find_drawing_file(rels_path, drawing_id)

            if drawing_file:
                image_mapping = self.parse_drawing(
                    os.path.join(temp_dir, 'xl/drawings', drawing_file),
                    os.path.join(temp_dir, 'xl/drawings/_rels', f'{drawing_file}.rels'),
                    image_dir
                )
        return image_mapping

    def find_drawing_file(self, rels_path, drawing_id):
        if os.path.exists(rels_path):
            trees = ET.parse(rels_path)
            tree=trees.getroot()
            return next(
                (rel.get('Target').split('/')[-1]
                 for rel in tree.findall('.//nsr:Relationship',self.namespaces)
                 if rel.get('Id') == drawing_id
                 ), None)
        return None

    def parse_drawing(self, drawing_path, rels_path, image_dir):
        rels = self.parse_relationships(rels_path)
        tree = ET.parse(drawing_path)

        image_data = {}
        for anchor in tree.findall('.//xdr:twoCellAnchor', self.namespaces):
            pic = anchor.find('.//xdr:pic', self.namespaces)
            if pic is not None:
                blip = pic.find('.//a:blip', self.namespaces)
                embed_id = blip.get('{%s}embed' % self.namespaces['r'])
                img_file = rels.get(embed_id)

                # 获取图片位置
                col = anchor.find('.//xdr:from/xdr:col', self.namespaces).text
                row = anchor.find('.//xdr:from/xdr:row', self.namespaces).text
                cell = f"{self.col_number_to_name(int(col))}{int(row) + 1}"

                # 保存图片并记录路径
                src_path = os.path.join(os.path.dirname(drawing_path), '..', 'media', img_file)
                dest_path = os.path.join(image_dir, img_file)
                shutil.copy(src_path, dest_path)

                image_data["关联图片"+cell] = dest_path
        return image_data

    def parse_relationships(self, rels_path):
        if not os.path.exists(rels_path):
            return {}
        else:
            tree = ET.parse(rels_path)
            return {rel.get('Id'): rel.get('Target').split('/')[-1]
                    for rel in tree.findall('.//nsr:Relationship',self.namespaces)}

    def col_number_to_name(self,n):
        # 将 n 调整为从1开始的索引，以便兼容原始算法
        n = n + 1
        name = ''
        while n > 0:
            n, rem = divmod(n - 1, 26)
            name = chr(65 + rem) + name
        return name

    def process_sheet_texts(self,root,xml_file,temp_dir):

        tree=ET.parse(os.path.join(root,xml_file))
        text_tree=ET.parse(os.path.join(temp_dir,'xl/sharedStrings.xml'))
        texts=text_tree.findall('.//ns:si',self.namespaces)
        cs=tree.findall('.//ns:sheetData/ns:row/ns:c',self.namespaces)
        text={}
        for c in cs:
            if c.get('t')=='s':
                location=c.get('r')
                if c.find('.//ns:v',self.namespaces) is not None:
                    value=c.find('.//ns:v',self.namespaces).text
                    if value.isdigit():
                        text[location]=texts[int(value)].find('.//ns:t',self.namespaces).text
                    else:
                        text[location]=value
        return text

    # def process_sheet(self, worksheet, sheet_images):
    #     sheet_datas = []
    #     sheet_data={}
    #     r=worksheet.merged_cells.ranges
    #     for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=2):
    #         if row[0] not in r:
    #             if row[0].value == None:
    #                 row[0].value="空"
    #             sheet_data['标准编号']=row[0].value
    #             sheet_data['标准标题']=row[1].value
    #             std_cell=row[1].offset(row=2)
    #             sheet_data['详细描述']=std_cell.value
    #             n_cell = std_cell.offset(row=1)
    #             while n_cell in r:
    #                 n_cell = n_cell.offset(row=1)
    #
    #             sheet_data['必要性']=n_cell.value
    #             break
    #     sheet_data['关联图片']=sheet_images
    #     sheet_datas.append(sheet_data)
    #     return sheet_datas

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel知识库构建工具")
        self.geometry("1000x600")
        self.processor = StructuredExcelProcessor()
        self.create_widgets()

    def create_widgets(self):
        # 文件选择部件
        self.file_frame = ttk.LabelFrame(self, text="Excel文件选择")
        self.file_frame.pack(pady=10, padx=10, fill=tk.X)

        self.file_entry = ttk.Entry(self.file_frame, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(self.file_frame, text="浏览", command=self.browse_file).pack(side=tk.LEFT)

        # 输出目录选择
        self.output_frame = ttk.LabelFrame(self, text="输出设置")
        self.output_frame.pack(pady=10, padx=10, fill=tk.X)

        ttk.Label(self.output_frame, text="输出目录:").pack(side=tk.LEFT)
        self.output_entry = ttk.Entry(self.output_frame, width=40)
        self.output_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(self.output_frame, text="浏览", command=self.browse_folder_output).pack(side=tk.LEFT)

        ttk.Label(self.output_frame, text="图片目录:").pack(side=tk.LEFT, padx=5)
        self.img_entry = ttk.Entry(self.output_frame, width=40)
        self.img_entry.pack(side=tk.LEFT)
        ttk.Button(self.output_frame, text="浏览", command=self.browse_folder_img).pack(side=tk.LEFT)

        # 日志输出
        self.log_frame = ttk.LabelFrame(self, text="处理日志")
        self.log_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(self.log_frame, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 进度条
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.pack(pady=5, padx=10, fill=tk.X)

        # 操作按钮
        self.btn_frame = ttk.Frame(self)
        self.btn_frame.pack(pady=10)

        ttk.Button(self.btn_frame, text="开始处理", command=self.start_processing).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.btn_frame, text="退出", command=self.quit).pack(side=tk.RIGHT, padx=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def browse_folder_output(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder_path)

    def browse_folder_img(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.img_entry.delete(0, tk.END)
            self.img_entry.insert(0, folder_path)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def start_processing(self):
        excel_path = self.file_entry.get()
        output_dir = self.output_entry.get()
        image_dir = self.img_entry.get()

        if not excel_path:
            self.log("错误：请先选择Excel文件")
            return

        try:
            self.progress['value'] = 0
            self.log("开始处理Excel文件...")

            # 执行处理
            self.processor.process_excel(
                excel_path,
                output_dir,
                image_dir
            )

            self.progress['value'] = 100
            self.log(f"处理完成！知识库已保存到{output_dir}目录")
        except Exception as e:
            error_message = traceback.format_exc()
            print(error_message)
            self.log(f"处理过程中发生错误：{str(e)}")


if __name__ == "__main__":
    app = Application()
    app.mainloop()