import tkinter as tk
import xlsxwriter
import calendar
import atexit
from tkinter import ttk, filedialog
from datetime import datetime

# Funkcija za generiranje Excel datoteke i unos podataka

workbook = xlsxwriter.Workbook('evidencija_radnog_vremena.xlsx')


def generiraj_evidenciju(dani_nerada):
    global worksheet, workbook, radni_list, file_path

    # Unos mjeseca i godine preko GUI-a
    mjesec = int(mjesec_entry.get())
    godina = int(godina_entry.get())
    radnik = radnik_entry.get()
    pocetak = int(pocetak_entry.get())
    kraj = int(kraj_entry.get())
    radni_list = radni_list_entry.get()

    broj_dana = calendar.monthrange(godina, mjesec)[1]

    worksheet_name = radni_list
    worksheets = {}
    worksheets[mjesec] = workbook.add_worksheet(worksheet_name)

    bold = workbook.add_format({'bold': True, 'align': 'center'})
    bold.set_text_wrap()
    regular = workbook.add_format({'align': 'center'})
    regular.set_text_wrap()
    border = workbook.add_format({'border': 2, 'align': 'center'})
    border.set_text_wrap()
    regular_left = workbook.add_format({'align': 'left', 'top': 2})
    regular_left.set_text_wrap()
    bottom_border = workbook.add_format({'bottom': 2, 'right': 2, 'indent': 2})
    bottom_border.set_text_wrap()
    right_border = workbook.add_format({'right': 2, 'align': 'center'})
    right_border.set_text_wrap()
    cell_format2 = workbook.add_format()
    cell_format2.set_indent(2)

    for red in range(5, 36):
        for stupac in range(2, 7):
            worksheets[mjesec].write_blank(red, stupac, '', border)

    worksheets[mjesec].merge_range('A1:G1', '', border)
    tekst = 'Evidencija o radnom vremenu za radnike koji poslove obavljaju na izdvojenom mjestu rada'
    worksheets[mjesec].write_string('A1', tekst, bold)

    def imena_mjeseci(mjesec):
        match mjesec:
            case 1:
                return "JANUAR"
            case 2:
                return "FEBRUAR"
            case 3:
                return "MART"
            case 4:
                return "APRIL"
            case 5:
                return "MAJ"
            case 6:
                return "JUNI"
            case 7:
                return "JULI"
            case 8:
                return "AUGUST"
            case 9:
                return "SEPTEMBAR"
            case 10:
                return "OKTOBAR"
            case 11:
                return "NOVEMBAR"
            case 12:
                return "DECEMBAR"
            case _:
                return "nepoznat mjesec"

    def praznici(mjesec):
        match mjesec:
            case 1:
                return worksheets[mjesec].write(5, 6, 8, border)
            case 3:
                return worksheets[mjesec].write(5, 6, 8, border)
            case 5:
                return worksheets[mjesec].write(5, 6, 8, border)
            case 11:
                return worksheets[mjesec].write(29, 6, 8, border)
            case 12:
                return worksheets[mjesec].write(29, 6, 8, border)

    worksheets[mjesec].merge_range('A2:G2', '', border)
    worksheets[mjesec].write_rich_string(
        'A2', "Za mjesec ", bold, imena_mjeseci(mjesec), regular, " godine ", bold, str(godina), border)

    worksheets[mjesec].merge_range('A3:G3', '', border)
    worksheets[mjesec].write_rich_string(
        'A3', bold, "Ime i prezime radnika: ", "__", radnik, "_____________", border)

    naslovi = ['R/B', 'Datum u mjesecu', 'Ukupno dnevno radno vrijeme',
               'Vrijeme terenskog rada', 'Vrijeme pripravnosti', 'Vrijeme neprisustva na poslu', 'Ukupno radnih sati']
    for i, naslov in enumerate(naslovi):
        worksheets[mjesec].write(3, i, naslov, border,)

    brojevi = list(range(1, 8))  # generira listu brojeva od 1 do 7
    for i, broj in enumerate(brojevi):
        worksheets[mjesec].write(4, i, broj, border)

    for i in range(1, 32):
        worksheets[mjesec].write(i + 4, 0, i, border)

    for dan in range(1, broj_dana + 1):
        datum = datetime(godina, mjesec, dan).strftime("%d.")
        worksheets[mjesec].write(5 + dan - 1, 1, datum, border)

    worksheets[mjesec].set_column('A:G', 80 / 7)

    pocetak_terenskog_rada = int(pocetak_terenskog_rada_entry.get())
    kraj_terenskog_rada = int(kraj_terenskog_rada_entry.get())

    for dan in range(1, broj_dana + 1):
        datum = datetime(godina, mjesec, dan)
        dan_u_tjednu = datum.weekday()
        datum_str = datum.strftime("%d.")
        worksheets[mjesec].write(5 + dan - 1, 1, datum_str, border)

        if pocetak_terenskog_rada <= dan <= kraj_terenskog_rada and dan_u_tjednu not in (5, 6):
            worksheets[mjesec].write(5 + dan - 1, 3, 8, border)

        dnevno_radno_vrijeme = 0
        if pocetak <= dan <= kraj and dan_u_tjednu not in (5, 6):
            dnevno_radno_vrijeme = 8
            worksheets[mjesec].write(
                5 + dan - 1, 2, dnevno_radno_vrijeme, border)
            worksheets[mjesec].write(
                5 + dan - 1, 6, dnevno_radno_vrijeme, border)
            praznici(mjesec)

        if dan_u_tjednu == 5 or dan_u_tjednu == 6:
            worksheets[mjesec].write(5 + dan - 1, 5, 2, border)

    dani_nerada = unesi_dane_nerada()
    for dan in range(1, broj_dana + 1):
        if dan in dani_nerada:

            # worksheets[mjesec].write_blank(5 + dan - 1, 6, '', border)
            worksheets[mjesec].write_blank(5 + dan - 1, 3, '', border)

    worksheets[mjesec].merge_range(
        36, 0, 36, 5, "Ukupno radnih sati u mjesecu:", border)
    worksheets[mjesec].write_formula('G37', '=SUM(G6:G36)', border)

    worksheets[mjesec].merge_range('A38:G38', '', right_border)
    worksheets[mjesec].write_string('A38', 'Napomena: ', regular_left)
    worksheets[mjesec].merge_range('A39:G39', '', right_border)
    worksheets[mjesec].write_string(
        'A39', 'U kolonu 6) Vrijeme neprisustva na poslu, potrebno je evidentirati vrijeme neprisustva i oznaku (broj) vrste neprisustva:', cell_format2)
    worksheets[mjesec].merge_range('A40:G40', '', right_border)
    worksheets[mjesec].write_string(
        'A40', ' 1- vrijeme korištenja odmora (sedmičnog i godišnjeg),', cell_format2)
    worksheets[mjesec].merge_range('A41:G41', '', right_border)
    worksheets[mjesec].write_string(
        'A41', ' 2- vrijeme za dane u kojima se ne radi i praznike utvrđene posebnim propisom, ', cell_format2)
    worksheets[mjesec].merge_range('A42:G42', '', right_border)
    worksheets[mjesec].write_string(
        'A42', ' 3- vrijeme spriječenosti za rad zbog privremene nesposobonosti za rad,', cell_format2)
    worksheets[mjesec].merge_range('A43:G43', '', right_border)
    worksheets[mjesec].write_string(
        'A43', ' 4- vrijeme porođajnog odsustva, roditeljskih dopusta, mirovanja radnog odnosa ili korištenja drugih prava u skladu s posebnim propisom,', cell_format2)
    worksheets[mjesec].merge_range('A44:G44', '', right_border)
    worksheets[mjesec].write_string(
        'A44', ' 5- vrijeme plaćenog odsustva,', cell_format2)
    worksheets[mjesec].merge_range('A45:G45', '', bottom_border)
    worksheets[mjesec].write_string(
        'A45', ' 6- vrijeme neplaćenog odsustva.', bottom_border)

    brisanje_entry()
    # ucitaj_podatke()

# Funkcija za unos dana kada radnik nije radio


def unesi_dane_nerada():
    dani_nerada = dani_nerada_entry.get()
    return [int(dan.strip()) for dan in dani_nerada.split(',') if dan.strip()]


# Funkcija za snimanje radnog lista pod prilagođenim imenom i nakon toga pravljenje novog radnog lista gdje se unose novi podaci
def brisanje_entry():
    radni_list_entry.delete(0, tk.END)
    radnik_entry.delete(0, tk.END)
    pocetak_entry.delete(0, tk.END)
    kraj_entry.delete(0, tk.END)
    pocetak_terenskog_rada_entry.delete(0, tk.END)
    kraj_terenskog_rada_entry.delete(0, tk.END)
    dani_nerada_entry.delete(0, tk.END)
    mjesec_entry.delete(0, tk.END)
    godina_entry.delete(0, tk.END)


def zatvori_workbook():
    global workbook
    if workbook is not None:
        workbook.close()


# def ucitaj_podatke():
#     path = "./evidencija_radnog_vremena.xlsx"
#     workbook = openpyxl.load_workbook(path)
#     sheet = workbook.active

#     list_values = list(sheet.values)
#     for col_name in list_values[3:2]:
#         treeview.heading(col_name, text=col_name)

#     for value_tuple in list_values[3:]:
#         treeview.insert('', tk.END, values=value_tuple)


atexit.register(zatvori_workbook)


# Glavni dio koda
if __name__ == "__main__":
    # Kreiranje tkinter prozora
    root = tk.Tk()
    style = ttk.Style(root)
    root.tk.call('source', 'forest-dark.tcl')
    style.theme_use('forest-dark')
    root.title("Evidencija radnog vremena")
    root_label_frame = ttk.LabelFrame(root, text="Evidencija radnog vremena")
    root_label_frame.pack(side="left", padx=20, pady=10, fill="y")

    # Unos radnog lista
    radni_list_label = ttk.Label(root_label_frame, text="Naziv radnog lista:")
    radni_list_label.pack()
    radni_list_entry = ttk.Entry(root_label_frame)
    radni_list_entry.pack()
    # Unos mjeseca
    mjesec_label = ttk.Label(root_label_frame, text="Mjesec (MM):")
    mjesec_label.pack()
    mjesec_entry = ttk.Entry(root_label_frame)
    mjesec_entry.pack()

    # Unos godine
    godina_label = ttk.Label(root_label_frame, text="Godina (GGGG):")
    godina_label.pack()
    godina_entry = ttk.Entry(root_label_frame)
    godina_entry.pack()

    # Unos radnika
    radnik_label = ttk.Label(root_label_frame, text="Ime i prezime radnika:")
    radnik_label.pack()
    radnik_entry = ttk.Entry(root_label_frame)
    radnik_entry.pack()

    # Kreiranje labela i polja za unos
    pocetak_label = ttk.Label(
        root_label_frame, text="Početak ukupnog dnevnog radnog vremena:")
    pocetak_entry = ttk.Entry(root_label_frame)
    kraj_label = ttk.Label(
        root_label_frame, text="Kraj ukupnog dnevnog radnog vremena:")
    kraj_entry = ttk.Entry(root_label_frame)

    # Pakiranje labela i polja za unos
    pocetak_label.pack()
    pocetak_entry.pack()
    kraj_label.pack()
    kraj_entry.pack()

    # Unos početnog dana terenskog rada
    pocetak_terenskog_rada_label = ttk.Label(
        root_label_frame, text="Prvi dan terenskog rada:")
    pocetak_terenskog_rada_label.pack()
    pocetak_terenskog_rada_entry = ttk.Entry(root_label_frame)
    pocetak_terenskog_rada_entry.pack()

    # Unos zadnjeg dana terenskog rada
    kraj_terenskog_rada_label = ttk.Label(
        root_label_frame, text="Zadnji dan terenskog rada:")
    kraj_terenskog_rada_label.pack()
    kraj_terenskog_rada_entry = ttk.Entry(root_label_frame)
    kraj_terenskog_rada_entry.pack()

    # Unos dana nerada
    dani_nerada_label = ttk.Label(
        root_label_frame, text="Dani nerada (odvojeni zarezom):")
    dani_nerada_label.pack()
    dani_nerada_entry = ttk.Entry(root_label_frame)
    dani_nerada_entry.pack()

    separator = ttk.Separator(root_label_frame)
    separator.pack(padx=(20, 10), pady=10, fill="x")

    # Dugme za generisanje evidencije
    generiraj_dugme = ttk.Button(root_label_frame, text="Generiraj evidenciju",
                                 command=lambda: generiraj_evidenciju(unesi_dane_nerada()))
    generiraj_dugme.pack(pady=10, side="bottom")

    # frame = ttk.Frame(root)
    # frame.pack(side="left", fill="both", expand=True)

    # treeFrame = ttk.Frame(frame)
    # treeFrame.grid(row=0, column=1, pady=10)
    # treeScroll = ttk.Scrollbar(treeFrame)
    # treeScroll.pack(side="right", fill="y")

    # cols = ("Ukupno dnevno radno vrijeme", "Vrijeme terenskog rada",
    #         "Vrijeme neprisustva na poslu", "Ukupno radnih sati")
    # treeview = ttk.Treeview(treeFrame, show="headings",
    #                         yscrollcommand=treeScroll.set, columns=cols, height=20)
    # treeview.column("Ukupno dnevno radno vrijeme", width=100)
    # treeview.column("Vrijeme terenskog rada", width=50)
    # treeview.column("Vrijeme neprisustva na poslu", width=100)
    # treeview.column("Ukupno radnih sati", width=100)
    # treeview.pack()
    # treeScroll.config(command=treeview.yview)

    # # Dugme za zatvaranje radnog lista
    # zatvori_workbook_dugme = ttk.Button(
    #     root, text="Zatvori radni list", command=zatvori_workbook)
    # zatvori_workbook_dugme.pack()

    # Pokretanje glavnog dijela programa
    root.mainloop()
