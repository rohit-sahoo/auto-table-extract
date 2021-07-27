import subprocess


def set_password(excel_file_path, pw):
    from pathlib import Path

    excel_file_path = Path(excel_file_path)

    vbs_script = \
        f"""' Save with password required upon opening

    Const xlSaveChanges = 1

    Set excel_object = CreateObject("Excel.Application")
    Set workbook = excel_object.Workbooks.Open("{excel_file_path}")

    excel_object.DisplayAlerts = False
    excel_object.Visible = False

    workbook.SaveAs "{excel_file_path}",, "{pw}"
    
    
    excel_object.Application.Quit
    """
    # write
    vbs_script_path = excel_file_path.parent.joinpath("set_pw.vbs")
    with open(vbs_script_path, "w") as file:
        file.write(vbs_script)
    # execute
    subprocess.call(['cscript.exe', str(vbs_script_path)])
    # remove
    vbs_script_path.unlink()

    return None
