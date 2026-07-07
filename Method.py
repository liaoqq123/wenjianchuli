import json
import shutil
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


DATA_FILE = Path(__file__).with_name("data.json")
DEFAULT_CONFIG = {
    "start_path": "",
    "target_path": "",
    "excel": "",
    "start": "1",
    "target": "2",
    "file_column": "1",
    "original_column": "2",
    "new_column": "3",
    "output_column": "4",
    "sheetname": "Sheet1",
}


class Method:
    @staticmethod
    def _safe_child_path(start_path, file_name):
        start_dir = Path(start_path).resolve()
        file_text = str(file_name).strip()
        if not file_text:
            return None

        relative_file = Path(file_text)
        if relative_file.is_absolute():
            return None

        child_path = (start_dir / relative_file).resolve()
        if not child_path.is_relative_to(start_dir):
            return None

        return child_path

    @staticmethod
    def getSelectedFiles(start_path, include_subdirs=False):
        start_dir = Path(start_path)
        if include_subdirs:
            return [path for path in start_dir.rglob("*") if path.is_file()]

        return [path for path in start_dir.iterdir() if path.is_file()]

    @staticmethod
    def getSelectedFolders(start_path, include_subdirs=False):
        start_dir = Path(start_path)
        if include_subdirs:
            return [path for path in start_dir.rglob("*") if path.is_dir()]

        return [path for path in start_dir.iterdir() if path.is_dir()]

    @staticmethod
    def _file_name_only(file_name):
        return Path(str(file_name).strip()).name

    @staticmethod
    def _read_name_pairs(excel_path, start, target, sheetname):
        workbook = load_workbook(excel_path)
        try:
            sheet = workbook[sheetname]
            start_column = int(start)
            target_column = int(target)
            name_pairs = []

            for line in range(2, sheet.max_row + 1):
                start_value = sheet.cell(line, start_column).value
                target_value = sheet.cell(line, target_column).value
                if start_value is None or target_value is None:
                    continue

                start_name = str(start_value).strip()
                target_name = str(target_value).strip()
                if start_name and target_name:
                    name_pairs.append((start_name, target_name))

            return name_pairs
        finally:
            workbook.close()

    @staticmethod
    def _rename_selected_paths(selected_paths, start_path, name_pairs, include_subdirs=False):
        renamed_paths = []
        start_dir = Path(start_path)

        for path in selected_paths:
            relative_name = str(path.relative_to(start_dir))
            relative_name_alt = relative_name.replace("\\", "/")

            for start_name, target_name in name_pairs:
                new_name = None

                if start_name in path.name:
                    new_name = path.name.replace(start_name, Method._file_name_only(target_name))
                elif include_subdirs and start_name in {relative_name, relative_name_alt}:
                    new_name = Method._file_name_only(target_name)

                if not new_name:
                    continue

                new_path = path.with_name(new_name)
                if new_path.exists():
                    continue

                path.rename(new_path)
                renamed_paths.append(str(new_path))
                break

        return renamed_paths

    @staticmethod
    def _find_files(start_path, file_name, include_subdirs=False):
        start_dir = Path(start_path).resolve()
        file_text = str(file_name).strip()
        if not file_text:
            return []

        direct_file = Method._safe_child_path(start_dir, file_text)
        if direct_file and direct_file.is_file():
            return [direct_file]

        if include_subdirs:
            return [path for path in Method.getSelectedFiles(start_dir, True) if path.name == file_text]

        return []

    @staticmethod
    def _copy_or_move(start_path, over_path, file_name, include_subdirs=False, move=False):
        start_dir = Path(start_path).resolve()
        target_dir = Path(over_path)
        handled_files = []

        for source_file in Method._find_files(start_path, file_name, include_subdirs):
            if include_subdirs:
                target_name = source_file.relative_to(start_dir)
            else:
                target_name = source_file.name

            target_file = target_dir / target_name
            if target_file.exists():
                continue

            target_file.parent.mkdir(parents=True, exist_ok=True)
            if move:
                shutil.move(str(source_file), str(target_file))
            else:
                shutil.copy2(source_file, target_file)

            handled_files.append(str(target_file))

        return handled_files

    # 文件写入方法
    @staticmethod
    def readMethod(
        start_path,
        excel_path,
        file_column,
        original_column,
        new_column,
        output_column,
        sheetname,
        include_subdirs=False,
    ):
        workbook = load_workbook(excel_path)
        try:
            sheet = workbook[sheetname]
            file_column = int(file_column)
            original_column = int(original_column)
            new_column = int(new_column)
            output_column = int(output_column)
            start_dir = Path(start_path)
            count = 0

            for column in {file_column, original_column, new_column, output_column}:
                for line in range(2, sheet.max_row + 1):
                    sheet.cell(line, column).value = None

            for row_index, file_path in enumerate(Method.getSelectedFiles(start_dir, include_subdirs), start=2):
                if include_subdirs:
                    file_name = str(file_path.relative_to(start_dir))
                else:
                    file_name = file_path.name

                original_name = file_path.stem
                new_name_cell = f"{get_column_letter(new_column)}{row_index}"

                sheet.cell(row_index, file_column).value = file_name
                sheet.cell(row_index, original_column).value = original_name
                sheet.cell(row_index, new_column).value = original_name
                sheet.cell(row_index, output_column).value = f'={new_name_cell}&"{file_path.suffix}"'
                count += 1

            workbook.save(excel_path)
            return count
        finally:
            workbook.close()

    # 初始化表格数据
    @staticmethod
    def initializeSheetMethod(excel_path, sheetname, file_column, original_column, new_column, output_column):
        workbook = load_workbook(excel_path)
        try:
            sheet = workbook[sheetname]
            file_column = int(file_column)
            original_column = int(original_column)
            new_column = int(new_column)
            output_column = int(output_column)

            for row in sheet.iter_rows():
                for cell in row:
                    cell.value = None

            sheet.cell(1, file_column).value = "文件名列"
            sheet.cell(1, original_column).value = "原名称列"
            sheet.cell(1, new_column).value = "新名称列"
            sheet.cell(1, output_column).value = "写入列"

            workbook.save(excel_path)
        finally:
            workbook.close()

    # 文件复制方法
    @staticmethod
    def copyMethod(start_path, over_path, file_name, include_subdirs=False):
        return Method._copy_or_move(start_path, over_path, file_name, include_subdirs, move=False)

    # 文件移动方法
    @staticmethod
    def moveMethod(start_path, over_path, file_name, include_subdirs=False):
        return Method._copy_or_move(start_path, over_path, file_name, include_subdirs, move=True)

    # 文件改名方法
    @staticmethod
    def renamingMethod(start_path, excel_path, start, target, sheetname, include_subdirs=False):
        name_pairs = Method._read_name_pairs(excel_path, start, target, sheetname)
        selected_files = Method.getSelectedFiles(start_path, include_subdirs)
        return Method._rename_selected_paths(selected_files, start_path, name_pairs, include_subdirs)

    # 文件夹改名方法
    @staticmethod
    def renamingFolderMethod(start_path, excel_path, start, target, sheetname, include_subdirs=False):
        name_pairs = Method._read_name_pairs(excel_path, start, target, sheetname)
        selected_folders = Method.getSelectedFolders(start_path, include_subdirs)
        selected_folders.sort(key=lambda path: len(path.parts), reverse=True)
        return Method._rename_selected_paths(selected_folders, start_path, name_pairs, include_subdirs)

    # 文件删除方法
    @staticmethod
    def removeMethod(start_path, file_name, include_subdirs=False):
        removed_files = []

        for file_path in Method._find_files(start_path, file_name, include_subdirs):
            file_path.unlink()
            removed_files.append(str(file_path))

        return removed_files

    # 读取json文件数据
    @staticmethod
    def readData():
        if not DATA_FILE.exists():
            return DEFAULT_CONFIG.copy()

        try:
            with DATA_FILE.open(mode="r", encoding="utf-8") as json_file:
                data = json.load(json_file)
        except (OSError, json.JSONDecodeError):
            return DEFAULT_CONFIG.copy()

        config = {**DEFAULT_CONFIG, **data}
        config["file_column"] = str(config.get("file_column") or config.get("start") or "1")
        config["original_column"] = str(config.get("original_column") or config.get("target") or "2")
        config["new_column"] = str(config.get("new_column") or "3")
        config["output_column"] = str(config.get("output_column") or "4")
        config["start"] = config["file_column"]
        config["target"] = config["new_column"]
        return config

    # 写入数据至json文件
    @staticmethod
    def storeData(data):
        with DATA_FILE.open("w", encoding="utf-8") as json_file:
            json.dump(data, json_file, ensure_ascii=False, indent=2)
