from tkinter import *
from tkinter.filedialog import askopenfilename
from pathlib import Path
import glob
import shutil
import os
import xlrd
import cv2
from tkinter import messagebox

AR, DE, EN, ES, FR, IT, NL, NO, PL, ES, PT, SE, RU, CN = "", "", "", "", "", "", "", "", "", "", "", "", "", ""
def set_url_values(val, index):
    global AR
    global DE
    global EN
    global FR
    global IT
    global NL
    global NO
    global PL
    global ES
    global PT
    global SE
    global RU
    global CN

    # print(val, index)
    if val == 'AR':
        cell1 = sheet.cell(index, 1)
        AR = cell1.value
    elif val == 'DE':
        cell2 = sheet.cell(index, 1)
        DE = cell2.value
    elif val == 'EN':
        cell3 = sheet.cell(index, 1)
        EN = cell3.value
    elif val == 'ES':
        cell4 = sheet.cell(index, 1)
        ES = cell4.value
    elif val == 'FR':
        cell5 = sheet.cell(index, 1)
        FR = cell5.value
    elif val == 'IT':
        cell6 = sheet.cell(index, 1)
        IT = cell6.value
    elif val == 'NL':
        cell7 = sheet.cell(index, 1)
        NL = cell7.value
    elif val == 'NO':
        cell8 = sheet.cell(index, 1)
        NO = cell8.value
    elif val == 'PL':
        cell9 = sheet.cell(index, 1)
        PL = cell9.value
    elif val == 'PT':
        cell10 = sheet.cell(index, 1)
        PT = cell10.value
    elif val == 'RU':
        cell11 = sheet.cell(index, 1)
        RU = cell11.value
    elif val == 'SV/SE':
        cell12 = sheet.cell(index, 1)
        SE = cell12.value
    elif val == 'CN':
        cell13 = sheet.cell(index, 1)
        CN = cell13.value
    else:
        print('Not Found', index, val)


def abc(nrows):
    for i in range(nrows):
        cell_val = sheet.cell(i, 0).value
        set_url_values(cell_val, i)


def update_html(folder_loc, html, find_str_list, replace_str_list, file):
    op = os.path.join(folder_loc + '/' + file + '.html')
    for find_str, replace_str in zip(find_str_list, replace_str_list):
        # print(html)
        html = html.replace(find_str, replace_str)
        html = html.replace(u'\xa0', u" ")
        # print(html)

    with open(op, 'w') as f:
        f.write(html)


def submit():
    global sheet
    src = src_dir.get()
    dst = dst_dir.get()

    src_dir.set("")
    dst_dir.set("")

    files = len([f for f in glob.glob(src + "**/*.jpg", recursive=True)])

    name = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xlsx'))])

    book = xlrd.open_workbook(name)
    sheet = book.sheet_by_index(0)
    num_rows = sheet.nrows
    book.release_resources()

    with open('html.html', 'r') as f:
        html = f.read()

    for jpgfile in glob.iglob(os.path.join(src, "*.jpg")):
        filename = os.path.basename(jpgfile)
        file, ext = os.path.splitext(filename)
        if ext == '.jpg' and file.endswith('-AR'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/ar')
            os.mkdir(dst + '/ar/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + AR + '"' + ' target="_blank">',
                                        '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-DE'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape
            abc(num_rows)
            output_file = os.path.join(dst + '/de')
            os.mkdir(dst + '/de/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + DE + '"' + ' target="_blank">',
                                '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)


        elif ext == '.jpg' and file.endswith('-EN'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/en')
            os.mkdir(dst + '/en/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + EN + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-ES'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/es')
            os.mkdir(dst + '/es/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + ES + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-FR'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/fr')
            os.mkdir(dst + '/fr/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + FR + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-IT'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/it')
            os.mkdir(dst + '/it/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + IT + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-NL'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/nl')
            os.mkdir(dst + '/nl/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + NL + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-NO'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/no')
            os.mkdir(dst + '/no/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + NO + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-Pl'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/pl')
            os.mkdir(dst + '/pl/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + PL + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-PT'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/pt')
            os.mkdir(dst + '/pt/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + PT + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-RU'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/ru')
            os.mkdir(dst + '/ru/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + RU + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-SV'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape

            abc(num_rows)
            output_file = os.path.join(dst + '/se')
            os.mkdir(dst + '/se/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + SE + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        elif ext == '.jpg' and file.endswith('-ZH'):
            src = os.path.join(src, jpgfile)
            img = cv2.imread(src)
            h, w, c = img.shape
            abc(num_rows)
            output_file = os.path.join(dst + '/cn')
            os.mkdir(dst + '/cn/')
            path = Path(output_file)
            shutil.copy(jpgfile, path)
            doc = os.path.join(output_file + '/' + file + '.html')
            find_list = ["<title></title>", "<tag1>", "<tag2>"]
            replace_list = ["<title>" + file + "</title>", '<a href="' + CN + '"' + ' target="_blank">',
                            '<img src="' + file + '.jpg"' + ' width="' + str(w) + '" height="' + str(h) + '" alt=""></a>']
            update_html(output_file, html, find_list, replace_list, file)

        else:
            continue

    # folder_counter = int(sum([len(folder) for _, _, folder in os.walk(dst)]) / 2)

    if files == num_rows:
        messagebox.showinfo("Successful", "All folders successfully created")
    else:
        messagebox.showwarning("Warning", "Folders Missing")


if __name__ == '__main__':

    top = Tk()
    top.geometry("450x300")
    top.iconbitmap('hnet.com-image.ico')
    top.title("efooter_generator")
    top.configure(background="Dark gray")

    input_path = Label(top, text="Enter Input Path:", bg="Dark gray").place(x=40, y=60)

    src_dir = StringVar()
    input_path_entry_area = Entry(top, textvariable=src_dir, width=30).place(x=150, y=60)

    output_path = Label(top, text="Enter Output Path:", bg="Dark gray").place(x=40, y=100)

    dst_dir = StringVar()
    output_path_entry_area = Entry(top, textvariable=dst_dir, width=30).place(x=150, y=100)

    submit_button = Button(top, text="Submit", command=submit).place(x=200, y=150)

    top.mainloop()








