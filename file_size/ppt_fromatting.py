import os
import glob
import win32com.client


# Folder path and file filter
ppt_folder = r"D:\Advance_Data_Sets\PPT_Sunset_NW"
ppt_files = glob.glob(os.path.join(ppt_folder, "UMTS_Sunset*.pptx"))

# Launch PowerPoint
ppt = win32com.client.Dispatch("PowerPoint.Application")
# Do not set ppt.Visible = False

for ppt_path in ppt_files:
    print(f"Processing: {ppt_path}")
    try:
        presentation = ppt.Presentations.Open(ppt_path, WithWindow=False)

        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.HasChart:
                    try:
                        chart = shape.Chart
                        workbook = chart.ChartData.Workbook
                        sheet = workbook.Worksheets(1)
                        used_range = sheet.UsedRange

                        for row in range(1, used_range.Rows.Count + 1):
                            for col in range(1, used_range.Columns.Count + 1):
                                value = sheet.Cells(row, col).Value
                                if value == 0:
                                    sheet.Cells(row, col).Value = None  # or "=NA()"

                    except Exception as chart_err:
                        print(f"  Chart error on Slide {slide.SlideIndex}: {chart_err}")

        presentation.Save()
        presentation.Close()

    except Exception as file_err:
        print(f"  Error opening {ppt_path}: {file_err}")

# Gracefully quit PowerPoint
try:
    if ppt:
        ppt.Quit()
except Exception as quit_err:
    print(f"⚠️ Could not quit PowerPoint: {quit_err}")

print("✅ All matching PowerPoint files processed.")
