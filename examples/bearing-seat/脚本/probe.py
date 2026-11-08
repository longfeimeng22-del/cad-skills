"""只读探针：确认可连接 AutoCAD 并读取活动图纸。"""
import win32com.client


def main():
    app = win32com.client.GetActiveObject("AutoCAD.Application")
    print("AutoCAD:", app.Version)
    count = app.Documents.Count
    print("已打开的文档数:", count)
    if count:
        doc = app.ActiveDocument
        print("活动文档:", doc.Name, "| 模型空间对象数:", doc.ModelSpace.Count)


if __name__ == "__main__":
    main()
