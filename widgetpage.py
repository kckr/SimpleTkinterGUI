import xlsxwriter
from xlsxwriter import worksheet
import time
from dbutil import DbUtil
from tkinter import ttk, messagebox
from tkinter import *
import print_pdf


class Gui:
    # widget_canvas = None
    # Declare global variable table1 record
    table1record = []
    table2record = []

    def __init__(self, root):  # initialize Gui
        self.root = root

        timestr = time.strftime("%Y%m%d-%H%M%S")
        DELETE_CHARS = '!@#$%^*=?\'\"{}[]<>!`:;|\\/,.'

        # delete specific characters
        def fix_string(s, old='', new=''):
            table = s.maketrans(old, new, DELETE_CHARS)
            return s.translate(table)

        root = Tk()

        root.title("BOM VIEW")
        root.geometry('870x800')

        cv = Canvas(root, height=950, width=950)
        cv.pack()
        scrollbar = Scrollbar(root)

        titlelable = Label(cv, text='', font=('Helvetica', 12))
        titlelable.grid(row=0, column=0)

        searchlable = Label(cv, text='Assembly Item Number', font=('Helvetica', 12))
        searchlable.grid(row=1, column=0)

        searchEntry = Entry(cv, width=37, font=('Helvetica', 12), highlightthickness=1, borderwidth=2,
                            bg='PaleGreen1')
        searchEntry.grid(row=3, column=0, ipady=3)

        assemblyitemlable = Label(cv, text='Assembly Item Description', font=('Helvetica', 12), width=40)
        assemblyitemlable.grid(row=1, column=2)

        assemblyitemdesc = ttk.Combobox(cv, width=41, textvariable='', font=('Helvetica', 10))
        assemblyitemdesc.grid(row=3, column=2, padx=10, sticky=E)
        assemblyitemdesc.selection_clear()

        readytextlabel = Label(cv, text='Ready', font=('Helvetica', 12))
        readytextlabel.grid(row=4, column=0, pady=10)

        readytext = ttk.Combobox(cv, width=39, values=('Yes', 'No'), font=('Helvetica', 11))
        readytext.grid(row=6, column=0, pady=8)
        readytext.selection_clear()

        shiptextlabel = Label(cv, text='Shipment', font=('Helvetica', 12), width=10)
        shiptextlabel.grid(row=4, column=2, pady=10)

        shiptext = ttk.Combobox(cv, width=37, values=('Yes', 'No'), font=('Helvetica', 11))
        shiptext.grid(row=6, column=2, pady=3, padx=5, sticky=E)
        shiptext.selection_clear()

        notetext = Label(cv, text='Note', font=('Helvetica', 12))
        notetext.grid(row=8, column=0, pady=10)

        notetextentry = Entry(cv, width=37, font=('Helvetica', 12), highlightthickness=2, borderwidth=2,
                              bg='LightCyan3')
        notetextentry.grid(row=9, column=0, pady=10)

        commenttext = Label(cv, text='Comments', font=('Helvetica', 12))
        commenttext.grid(row=10, column=0, pady=5)

        commenttextentry = Entry(cv, width=37, font=('Helvetica', 12), highlightthickness=2, borderwidth=2,
                                 bg='LightCyan3')
        commenttextentry.grid(row=11, column=0, pady=5, ipady=7)

        aftercommentlabel = Label(cv, text='', font=('Helvetica', 12))
        aftercommentlabel.grid(row=12, columnspan=3)

        def fixed_map(option):
            # Fix for setting text colour for Tkinter 8.6.9
            # From: https://core.tcl.tk/tk/info/509cafafae
            #
            # Returns the style map for 'option' with any styles starting with
            # ('!disabled', '!selected', ...) filtered out.

            # style.map() returns an empty list for missing options, so this
            # should be future-safe.
            return [elm for elm in style.map('Treeview', query_opt=option) if
                    elm[:2] != ('!disabled', '!selected')]

        style = ttk.Style()
        style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))

        viewtable = ttk.Treeview(cv, height=18, )

        viewtable.grid(row=13, columnspan=4, padx=41, sticky=W)
        viewtable.config(yscrollcommand=scrollbar.set)

        scrollbar.config(command=viewtable.yview)
        scrollbar.place(in_=viewtable, relx=1, y=1, relheight=1, bordermode="outside")

        style.configure(".", font=('Helvetica', 10), width=60)
        style.configure("Treeview.Heading", foreground='green', background='yellow', font=('Helvetica', 12))

        viewtable["columns"] = ["COMPONENT DESCRIPTION", "COMPONENT NUMBER", "QUANTITY"]
        viewtable["show"] = "headings"
        viewtable.heading("COMPONENT DESCRIPTION", text="COMPONENT DESCRIPTION")
        viewtable.column("COMPONENT DESCRIPTION", width=320)
        viewtable.heading("COMPONENT NUMBER", text="COMPONENT NUMBER ")
        viewtable.heading("QUANTITY", text="QUANTITY")

        def clicked():
            if not searchEntry.get():
                messagebox.showinfo("Bom Viewer", 'No input Detected')
            else:
                viewtable.delete(*viewtable.get_children())
                data = DbUtil('XML', 'COMPONENT_ROW', 'ASSEMBLY_ITEM')
                itemdescription = data.readdata(searchEntry.get())
                self.table1record = itemdescription

                cpt = 0

                for itemno, componentno, desc, qty in itemdescription:  # from 1st record
                    viewtable.tag_configure('oddrow', background='lavender')

                    viewtable.tag_configure('evenrow', background='alice blue')

                    if cpt % 2 == 0:
                        viewtable.insert('', 'end', text=str(cpt),
                                         values=(desc, componentno, qty), tags=('evenrow',))
                    else:
                        viewtable.insert('', 'end', text=str(cpt),
                                         values=(desc, componentno, qty), tags=('oddrow',))
                    cpt += 1

                table2data = data.readTable2(searchEntry.get())
                self.table2record = str(table2data)

                newstr = fix_string(str(table2data))

                assemblyitemdesc['values'] = newstr
                assemblyitemdesc.set(newstr)

        btn = Button(cv, text='Search', font=('Helvetica', 10, 'bold'), foreground='white',
                     background='cornflower blue',
                     highlightthickness=2.5, borderwidth=2.5, command=clicked)
        btn.grid(row=3, column=1, padx=10)

        btn.place(in_=searchEntry, relx=0.98, y=-1, bordermode="outside")  # keep widgets side by side

        def printtopdf():
            if not searchEntry.get():
                messagebox.showinfo("Bom Viewer", 'No input Detected')
            else:
                print_pdf.text((str(searchEntry.get()) + timestr + """.pdf"""), searchEntry.get(),
                               fix_string(str(self.table2record)), readytext.get(), shiptext.get(), self.table1record,
                               notetextentry.get(), commenttextentry.get())

        btn1 = Button(cv, text='Save_pdf', width=10, font=('Courier', 14, 'bold'), foreground='white',
                      background='medium sea green',
                      command=printtopdf)

        btn1.place(in_=commenttextentry, relx=1, y=-0.4, bordermode="outside")

        def where_used():
            window = Toplevel()
            window.title("Where Used")
            root.geometry('870x800')

            canvas1 = Canvas(window, height=750, width=750)
            canvas1.pack()

            scrollbar = Scrollbar(canvas1)

            titlelable = Label(canvas1, text='', font=('Helvetica', 12))
            titlelable.grid(row=0, column=0)

            componentLable = Label(canvas1, text='Component Number', font=('Helvetica', 12))
            componentLable.grid(row=1, column=0, padx=15)

            componentEntry = Entry(canvas1, width=30, font=('Helvetica', 11), highlightthickness=1, borderwidth=1,
                                   bg='PaleGreen1')
            componentEntry.grid(row=3, column=0, ipady=1)

            componentDesclable = Label(canvas1, text='Component Description', font=('Helvetica', 12), width=35)
            componentDesclable.grid(row=1, column=2)

            componentDesc = Entry(canvas1, width=37, font=('Helvetica', 10), highlightthickness=1, borderwidth=1,
                                  bg='PaleGreen1')
            componentDesc.grid(row=3, column=2, ipady=1)

            blanklable = Label(canvas1, text='', font=('Helvetica', 12))
            blanklable.grid(row=4, column=0)

            # Apply Style to treeView
            style = ttk.Style()
            style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))

            viewtable = ttk.Treeview(canvas1, height=28)
            viewtable.grid(row=5, columnspan=3, padx=31, sticky=NSEW)
            viewtable.config(yscrollcommand=scrollbar.set)

            scrollbar.config(command=viewtable.yview)
            scrollbar.place(in_=viewtable, relx=1, y=1, relheight=1, bordermode="outside")

            style.configure(".", font=('Helvetica', 10), width=60)
            style.configure("Treeview.Heading", foreground='green', background='yellow', font=('Helvetica', 10))

            # define columns strings
            viewtable["columns"] = ["ASSEMBLY ITEM_NUMBER", "ASSEMBLY ITEM_DESCRIPTION"]
            # show column heading
            viewtable["show"] = "headings"
            viewtable.heading("ASSEMBLY ITEM_NUMBER", text="ASSEMBLY ITEM_NUMBER")
            viewtable.column("ASSEMBLY ITEM_NUMBER")
            viewtable.heading("ASSEMBLY ITEM_DESCRIPTION", text="ASSEMBLY ITEM_DESCRIPTION")

            def where_used_clicked():  # component search function
                if not componentEntry.get():
                    messagebox.showinfo('Component Search', 'No Input')
                else:
                    viewtable.delete(*viewtable.get_children())
                    data = DbUtil('XML', 'COMPONENT_ROW', 'ASSEMBLY_ITEM')
                    component_data = data.read_component(componentEntry.get())
                    print(str(component_data))

                    cpt = 0
                    for ids, comp_no in component_data:
                        descdata = data.read_description(ids)

                        a = fix_string(str(descdata))
                        componentDesc.delete(0, END)  # insert in Entry widget
                        componentDesc.insert(0, comp_no)

                        viewtable.tag_configure('oddrow', background='lavender')

                        viewtable.tag_configure('evenrow', background='alice blue')

                        if cpt % 2 == 0:
                            viewtable.insert('', 'end', text=str(cpt),
                                             values=(ids, a), tags=('evenrow',))
                        else:
                            viewtable.insert('', 'end', text=str(cpt),
                                             values=(ids, a), tags=('oddrow',))
                        cpt += 1

                    def export_to_xls():  # excel file export function
                        # define workbook and add worksheet
                        workbook = xlsxwriter.Workbook(componentEntry.get() + timestr + """.xlsx """)
                        worksheet = workbook.add_worksheet()
                        row = 1
                        col = 0
                        bold = workbook.add_format({'bold': True})

                        worksheet.write('A1', 'Component Number', bold)
                        worksheet.set_column('A:A', 30, None)

                        worksheet.write('B1', 'Component Description', bold)
                        worksheet.set_column('B:B', 35, None)

                        worksheet.write('C1', 'Assembly Item Number', bold)
                        worksheet.set_column('C:C', 30, None)

                        worksheet.write('D1', 'Assembly Item Description', bold)
                        worksheet.set_column('D:D', 44, None)

                        # insert data to excel
                        for assemids, comp_num in component_data:
                            descdata = data.read_description(assemids)

                            a = fix_string(str(descdata))
                            worksheet.write(row, col, componentEntry.get())
                            worksheet.write(row, col + 1, componentDesc.get())
                            worksheet.write(row, col + 2, assemids)
                            worksheet.write(row, col + 3, a)
                            row += 1  # auto row increament

                        workbook.close()  # writes into excel

                    exportbtn = Button(canvas1, text='Export to excel', font=('Helvetica', 10, 'bold'),
                                       foreground='white',
                                       background='medium sea green',
                                       highlightthickness=2, borderwidth=2, width=15, command=export_to_xls)
                    exportbtn.grid(row=4, column=2)

            search_btn = Button(canvas1, text='Search', font=('Helvetica', 10, 'bold'), foreground='white',
                                background='cornflower blue',
                                highlightthickness=2, borderwidth=2, width=7, command=where_used_clicked)
            search_btn.grid(row=3, column=1, padx=0, sticky=W)

        # clear the screen
        def clear():
            viewtable.delete(*viewtable.get_children())
            searchEntry.delete(0, 'end')
            assemblyitemdesc.delete(0, 'end')
            shiptext.delete(0, 'end')
            readytext.delete(0, 'end')
            notetextentry.delete(0, 'end')
            commenttextentry.delete(0, 'end')

        btn3 = Button(cv, text='Clear', width=8, font=('Courier', 14, 'bold'), foreground='white',
                      background='pale violet red', command=clear)
        btn3.place(in_=btn1, relx=1, y=-0.4, bordermode="outside")

        btn2 = Button(cv, text='Component Search', width=17, font=('Courier', 14, 'bold'), foreground='white',
                      background='salmon3', command=where_used)
        btn2.place(in_=btn3, relx=1, y=-0.4, bordermode="outside")

        root.mainloop()


Gui(ttk)
