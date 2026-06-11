# -*- coding: utf-8 -*-
"""
PDF til Word - dra inn en PDF, få en Word-fil ut.
All behandling skjer lokalt. Ingen valg, ingen oppsett.
"""

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND = True
except ImportError:
    DND = False

import converter

BLAA = "#2563eb"
GRAA = "#6b7280"
GRONN = "#16a34a"
ROED = "#dc2626"


class App((TkinterDnD.Tk if DND else tk.Tk)):
    def __init__(self):
        super().__init__()
        self.title("PDF_To_Word_Converter")
        self.geometry("480x320")
        self.minsize(420, 280)
        self.configure(bg="white")
        self.jobber = False
        self.siste_utfil = None
        self._bygg_ui()

        if DND:
            self.drop_zone.drop_target_register(DND_FILES)
            self.drop_zone.dnd_bind("<<Drop>>", self._paa_slipp)

    def _bygg_ui(self):
        self.drop_zone = tk.Frame(self, bg="#f8fafc", highlightthickness=2,
                                  highlightbackground="#cbd5e1")
        self.drop_zone.pack(fill="both", expand=True, padx=14, pady=14)
        self.drop_zone.bind("<Button-1>", lambda e: self._velg_fil())

        self.ikon = tk.Label(self.drop_zone, text="\u2913", font=("Segoe UI", 34),
                             bg="#f8fafc", fg=BLAA)
        self.ikon.pack(pady=(34, 0))

        tekst = "Slipp en PDF her" if DND else "Klikk for å velge PDF"
        self.tittel = tk.Label(self.drop_zone, text=tekst,
                               font=("Segoe UI", 14, "bold"), bg="#f8fafc", fg="#111")
        self.tittel.pack()

        under = "eller klikk for å velge fil" if DND else ""
        self.under = tk.Label(self.drop_zone, text=under, font=("Segoe UI", 10),
                              bg="#f8fafc", fg=GRAA)
        self.under.pack()

        self.status = tk.Label(self.drop_zone, text="Word-filen lagres i samme mappe som PDF-en.",
                               font=("Segoe UI", 9), bg="#f8fafc", fg=GRAA,
                               wraplength=400, justify="center")
        self.status.pack(pady=(10, 0))

        self.fremdrift = ttk.Progressbar(self.drop_zone, maximum=100, length=300)

        self.aapne_knapp = tk.Button(self.drop_zone, text="Vis filen",
                                     command=self._aapne_mappe, relief="flat",
                                     bg=BLAA, fg="white", padx=14, pady=4,
                                     activebackground="#1d4ed8",
                                     activeforeground="white", cursor="hand2")

        for w in (self.ikon, self.tittel, self.under, self.status):
            w.bind("<Button-1>", lambda e: self._velg_fil())

    # --- hendelser ---

    def _paa_slipp(self, event):
        stier = self.tk.splitlist(event.data)
        if stier:
            self._start(stier[0])

    def _velg_fil(self):
        if self.jobber:
            return
        sti = filedialog.askopenfilename(
            title="Velg PDF-fil",
            filetypes=[("PDF-filer", "*.pdf")])
        if sti:
            self._start(sti)

    def _start(self, pdf):
        if self.jobber:
            return
        if not pdf.lower().endswith(".pdf"):
            self._sett_status("Filen må være en PDF.", ROED)
            return
        self.jobber = True
        self.siste_utfil = None
        self.aapne_knapp.pack_forget()
        self.fremdrift["value"] = 0
        self.fremdrift.pack(pady=(10, 0))
        self.tittel.config(text=os.path.basename(pdf))
        self.under.config(text="")
        threading.Thread(target=self._kjor, args=(pdf,), daemon=True).start()

    def _kjor(self, pdf):
        try:
            res = converter.konverter(
                pdf,
                status=lambda s: self.after(0, self._sett_status, s, GRAA),
                fremdrift=lambda p: self.after(0, self.fremdrift.config, {"value": p}),
            )
            self.siste_utfil = res["utfil"]
            navn = os.path.basename(res["utfil"])
            ekstra = " (med tekstgjenkjenning)" if res["ocr_brukt"] else ""
            self.after(0, self._ferdig, f"Ferdig{ekstra}: {navn}")
        except Exception as e:
            self.after(0, self._feil, str(e))

    def _ferdig(self, melding):
        self.jobber = False
        self.fremdrift.pack_forget()
        self._sett_status(melding, GRONN)
        self.tittel.config(text="Slipp en ny PDF her" if DND else "Klikk for å velge ny PDF")
        self.aapne_knapp.pack(pady=(10, 0))

    def _feil(self, melding):
        self.jobber = False
        self.fremdrift.pack_forget()
        self._sett_status(f"Noe gikk galt: {melding}", ROED)
        self.tittel.config(text="Prøv igjen med en annen PDF")

    def _sett_status(self, tekst, farge):
        self.status.config(text=tekst, fg=farge)

    def _aapne_mappe(self):
        if not self.siste_utfil:
            return
        mappe = os.path.dirname(self.siste_utfil)
        if sys.platform == "win32":
            subprocess.Popen(["explorer", "/select,", self.siste_utfil])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", self.siste_utfil])
        else:
            subprocess.Popen(["xdg-open", mappe])


if __name__ == "__main__":
    App().mainloop()
