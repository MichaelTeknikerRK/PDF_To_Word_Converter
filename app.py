# -*- coding: utf-8 -*-
"""
PDF_To_Word_Converter - dra inn en PDF, få en Word-fil ut.
All behandling skjer lokalt på maskinen.
"""

import multiprocessing
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog

import customtkinter as ctk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND = True
except ImportError:
    DND = False

import converter

BLAA = "#2563eb"
BLAA_MORK = "#1d4ed8"
GRAA = "#64748b"
GRONN = "#16a34a"
ROED = "#dc2626"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class App(ctk.CTk, *( [TkinterDnD.DnDWrapper] if DND else [] )):
    def __init__(self):
        super().__init__()
        if DND:
            self.TkdndVersion = TkinterDnD._require(self)

        self.title("PDF_To_Word_Converter")
        self.geometry("560x460")
        self.minsize(500, 420)
        self.configure(fg_color="#eef2f7")

        self.prosess = None
        self.ko = None
        self.siste_utfil = None
        self.fjern_marger = ctk.BooleanVar(value=True)

        self._bygg_ui()

        if DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._paa_slipp)

    # ------------------------------------------------------------ UI

    def _bygg_ui(self):
        # Toppfelt
        topp = ctk.CTkFrame(self, fg_color="transparent")
        topp.pack(fill="x", padx=24, pady=(20, 0))
        ctk.CTkLabel(topp, text="PDF til Word",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color="#0f172a").pack(side="left")
        ctk.CTkLabel(topp, text="Behandles lokalt - ingenting lastes opp",
                     font=ctk.CTkFont(size=12), text_color=GRAA).pack(side="right")

        # Slippsone (kort)
        self.kort = ctk.CTkFrame(self, corner_radius=16, fg_color="white",
                                 border_width=2, border_color="#dbe3ee")
        self.kort.pack(fill="both", expand=True, padx=24, pady=16)

        self.ikon = ctk.CTkLabel(self.kort, text="\u2913",
                                 font=ctk.CTkFont(size=44), text_color=BLAA)
        self.ikon.pack(pady=(46, 4))

        self.tittel = ctk.CTkLabel(
            self.kort, text="Slipp en PDF her" if DND else "Klikk for å velge PDF",
            font=ctk.CTkFont(size=17, weight="bold"), text_color="#0f172a")
        self.tittel.pack()

        self.under = ctk.CTkLabel(
            self.kort, text="eller klikk for å velge fil" if DND else "",
            font=ctk.CTkFont(size=12), text_color=GRAA)
        self.under.pack(pady=(2, 8))

        self.status = ctk.CTkLabel(self.kort,
                                   text="Word-filen lagres i samme mappe som PDF-en",
                                   font=ctk.CTkFont(size=12), text_color=GRAA,
                                   wraplength=420)
        self.status.pack()

        self.fremdrift = ctk.CTkProgressBar(self.kort, width=340, height=8,
                                            progress_color=BLAA)
        self.fremdrift.set(0)

        knapper = ctk.CTkFrame(self.kort, fg_color="transparent")
        knapper.pack(side="bottom", pady=(0, 18))

        self.avbryt_knapp = ctk.CTkButton(
            knapper, text="Avbryt", width=110, fg_color="#e2e8f0",
            text_color="#0f172a", hover_color="#cbd5e1",
            command=self._avbryt)

        self.vis_knapp = ctk.CTkButton(
            knapper, text="Vis filen i mappen", width=160,
            fg_color=BLAA, hover_color=BLAA_MORK, command=self._vis_fil)

        # Klikk hvor som helst på kortet for å velge fil
        for w in (self.kort, self.ikon, self.tittel, self.under, self.status):
            w.bind("<Button-1>", lambda e: self._velg_fil())

        # Bunnlinje med innstilling
        bunn = ctk.CTkFrame(self, fg_color="transparent")
        bunn.pack(fill="x", padx=24, pady=(0, 16))
        ctk.CTkCheckBox(bunn, text="Fjern topp- og bunntekst (gjentatte linjer og sidetall)",
                        variable=self.fjern_marger,
                        font=ctk.CTkFont(size=12), text_color="#334155",
                        checkbox_width=18, checkbox_height=18,
                        border_width=2).pack(side="left")

    # ------------------------------------------------------- hendelser

    def _paa_slipp(self, event):
        stier = self.tk.splitlist(event.data)
        if stier:
            self._start(stier[0])

    def _velg_fil(self):
        if self.prosess:
            return
        sti = filedialog.askopenfilename(title="Velg PDF-fil",
                                         filetypes=[("PDF-filer", "*.pdf")])
        if sti:
            self._start(sti)

    def _start(self, pdf):
        if self.prosess:
            return
        if not pdf.lower().endswith(".pdf"):
            self._sett_status("Filen må være en PDF.", ROED)
            return

        self.siste_utfil = None
        self.vis_knapp.pack_forget()
        self.tittel.configure(text=os.path.basename(pdf))
        self.under.configure(text="")
        self.fremdrift.set(0)
        self.fremdrift.pack(pady=(14, 0))
        self.avbryt_knapp.pack()

        self.ko = multiprocessing.Queue()
        self.prosess = multiprocessing.Process(
            target=converter.kjor_i_prosess,
            args=(self.ko, pdf, self.fjern_marger.get()),
            daemon=True)
        self.prosess.start()
        self.after(150, self._poll)

    def _poll(self):
        if not self.prosess:
            return
        try:
            while True:
                type_, data = self.ko.get_nowait()
                if type_ == "status":
                    self._sett_status(data, GRAA)
                elif type_ == "fremdrift":
                    self.fremdrift.set(data / 100)
                elif type_ == "ferdig":
                    self._ferdig(data)
                    return
                elif type_ == "feil":
                    self._feil(data)
                    return
        except Exception:
            pass  # køen er tom

        # Døde prosessen uten beskjed? (minnekrasj o.l.)
        if not self.prosess.is_alive():
            self._feil("Konverteringen stoppet uventet. Dokumentet kan være for "
                       "stort eller komplekst - prøv å dele det opp i mindre deler.")
            return
        self.after(150, self._poll)

    def _ferdig(self, res):
        self._rydd_prosess()
        self.siste_utfil = res["utfil"]
        deler = []
        if res.get("ocr_brukt"):
            deler.append("tekstgjenkjenning brukt")
        if res.get("fjernet_blokker"):
            deler.append(f"{res['fjernet_blokker']} topp/bunn-elementer fjernet")
        ekstra = f" ({', '.join(deler)})" if deler else ""
        self._sett_status(f"Ferdig{ekstra}: {os.path.basename(res['utfil'])}", GRONN)
        self.tittel.configure(text="Slipp en ny PDF her" if DND else "Klikk for ny PDF")
        self.vis_knapp.pack()

    def _feil(self, melding):
        self._rydd_prosess()
        self._sett_status(f"Noe gikk galt: {melding}", ROED)
        self.tittel.configure(text="Prøv igjen")

    def _avbryt(self):
        if self.prosess and self.prosess.is_alive():
            self.prosess.terminate()
        self._rydd_prosess()
        self._sett_status("Avbrutt.", GRAA)
        self.tittel.configure(text="Slipp en PDF her" if DND else "Klikk for å velge PDF")

    def _rydd_prosess(self):
        self.prosess = None
        self.ko = None
        self.fremdrift.pack_forget()
        self.avbryt_knapp.pack_forget()

    def _sett_status(self, tekst, farge):
        self.status.configure(text=tekst, text_color=farge)

    def _vis_fil(self):
        if not self.siste_utfil:
            return
        sti = os.path.normpath(self.siste_utfil)  # Explorer krever omvendte skråstreker
        if sys.platform == "win32":
            subprocess.Popen(f'explorer /select,"{sti}"')
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", sti])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(sti)])


if __name__ == "__main__":
    multiprocessing.freeze_support()  # KRITISK for PyInstaller + multiprocessing
    App().mainloop()
