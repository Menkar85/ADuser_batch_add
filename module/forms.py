import gettext
import pickle
import tkinter as tk
from tkinter import messagebox, font, filedialog

from application import LANGUAGE_CODE, LAST_VALUES, LOCALE_DIR

LANGUAGE_CODE = LANGUAGE_CODE


def set_language(code):
    global LANGUAGE_CODE
    global _
    LANGUAGE_CODE = code
    lang = gettext.translation('application', localedir=LOCALE_DIR, languages=[code], fallback=True)
    lang.install()
    _ = lang.gettext


# Load default translations
set_language(LANGUAGE_CODE)


class InputForm:
    """
    Class defining variables and structure of the graphical interface.
    """

    def __init__(self, master):
        self.master = master
        self.master.title("Mass input of users.")

        # initialize variables
        self.ldap_server = tk.StringVar()
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.source_file = tk.StringVar()
        self.destination_ou = tk.StringVar()
        self.domain = tk.StringVar()
        self.upn_suffix = tk.StringVar()
        self.result_file = tk.StringVar()
        self.logfile = tk.StringVar()

        # Load last values

        self._load_last_values()

        # Fonts
        label_font = font.Font(family='Times New Roman', size=14, weight='bold')
        entry_font = font.Font(family='Times New Roman', size=14)

        # labels and entries
        label_servername = tk.Label(master, text="AD Server name:", font=label_font)
        label_servername.grid(sticky='e', row=0, column=0, padx=(30, 5), pady=10)
        entry_servername = tk.Entry(master, width=40, textvariable=self.ldap_server, font=entry_font)
        entry_servername.grid(row=0, column=1, padx=(5, 30), pady=10, ipady=3)

        label_username = tk.Label(master, text="Username:", font=label_font)
        label_username.grid(sticky='e', row=1, column=0, padx=(30, 5), pady=10)
        entry_username = tk.Entry(master, width=40, textvariable=self.username, font=entry_font)
        entry_username.grid(row=1, column=1, padx=(5, 30), pady=10, ipady=3)

        label_password = tk.Label(master, text="Password:", font=label_font)
        label_password.grid(sticky='e', row=2, column=0, padx=(30, 5), pady=10)
        entry_password = tk.Entry(master, show='*', width=40, textvariable=self.password,
                                  font=entry_font)  # Password should be masked
        entry_password.grid(row=2, column=1, padx=(5, 30), pady=10, ipady=3)

        label_source_filename = tk.Label(master, text="Source filename:", font=label_font)
        label_source_filename.grid(sticky='e', row=3, column=0, padx=(30, 5), pady=10)
        entry_source_filename = tk.Entry(master, width=36, textvariable=self.source_file, font=entry_font)
        entry_source_filename.grid(row=3, sticky='w', column=1, padx=(5, 30), pady=10, ipady=3)

        label_destination_ou = tk.Label(master, text="Destination OU:", font=label_font)
        label_destination_ou.grid(sticky='e', row=4, column=0, padx=(30, 5), pady=10)
        entry_destination_ou = tk.Entry(master, width=40, textvariable=self.destination_ou, font=entry_font)
        entry_destination_ou.grid(row=4, column=1, padx=(5, 30), pady=10, ipady=3)

        label_domain = tk.Label(master, text="Domain:", font=label_font)
        label_domain.grid(sticky='e', row=5, column=0, padx=(30, 5), pady=10)
        entry_domain = tk.Entry(master, width=40, textvariable=self.domain, font=entry_font)
        entry_domain.grid(row=5, column=1, padx=(5, 30), pady=10, ipady=3)

        label_upn = tk.Label(master, text="UPN-suffix:", font=label_font)
        label_upn.grid(sticky='e', row=6, column=0, padx=(30, 5), pady=10)
        entry_upn = tk.Entry(master, width=40, textvariable=self.upn_suffix, font=entry_font)
        entry_upn.grid(row=6, column=1, padx=(5, 30), pady=10, ipady=3)

        label_result_file = tk.Label(master, text="Results file name:", font=label_font)
        label_result_file.grid(sticky='e', row=7, column=0, padx=(30, 5), pady=10)
        entry_result_file = tk.Entry(master, width=40, textvariable=self.result_file, font=entry_font)
        entry_result_file.grid(row=7, column=1, padx=(5, 30), pady=10, ipady=3)

        label_logfile = tk.Label(master, text="Log file name:", font=label_font)
        label_logfile.grid(sticky='e', row=8, column=0, padx=(30, 5), pady=10)
        entry_logfile = tk.Entry(master, width=40, textvariable=self.logfile, font=entry_font)
        entry_logfile.grid(row=8, column=1, padx=(5, 30), pady=10, ipady=3)

        # buttons

        browse_button = tk.Button(master, text='\U0001F4C2', command=self._browse_filename,
                                  font=font.Font(family='Times New Roman', size=12))
        browse_button.grid(row=3, column=1, sticky='e', padx=(0, 30))

        button_start = tk.Button(master, text="Start import", command=self._on_start, font=label_font)
        button_start.grid(row=9, columnspan=2, column=0, padx=50, pady=15, ipadx=129)

        button_cancel = tk.Button(master, text="Cancel", command=self._on_cancel, font=label_font)
        button_cancel.grid(row=10, columnspan=2, column=0, padx=50, pady=15, ipadx=150)

        # Language Selection Dropdown
        self.language_var = tk.StringVar(value=LANGUAGE_CODE)
        self.language_var.trace_add('write', lambda *args: self.update_language(self.language_var.get()))
        language_menu = tk.OptionMenu(self.master, self.language_var, "en_US", "ru_RU")
        language_menu.grid(row=11, column=0, padx=10, pady=5)

    def _on_start(self):
        """
        Requirement for confirmation of import
        Returns: None

        """
        # Show a confirmation dialog before proceeding
        confirm = messagebox.askyesno(_("Confirmation"), _(f"Do you want to start import?\n"))

        if confirm:
            self._save_last_values()
            self.master.destroy()  # Close the form window

    def _on_cancel(self):
        """
        Close window and exit program.
        Returns:

        """
        # Show a confirmation dialog before cancelling
        confirm = messagebox.askyesno(_("Confirmation"), _("Are you sure you want to cancel?"))

        if confirm:
            self.master.destroy()
            exit()  # Close the form window

    def _save_last_values(self):
        values = {
            'ldap_server': self.ldap_server.get(),
            'username': self.username.get(),
            'source_file': self.source_file.get(),
            'destination_ou': self.destination_ou.get(),
            'domain': self.domain.get(),
            'upn_suffix': self.upn_suffix.get(),
            'result_file': self.result_file.get(),
            'logfile': self.logfile.get(),
        }
        with open(LAST_VALUES, 'wb') as fp:
            pickle.dump(values, fp)

    def _load_last_values(self):
        try:
            with open(LAST_VALUES, 'rb') as fp:
                values = pickle.load(fp)
                self.ldap_server.set(values['ldap_server'])
                self.username.set(values['username'])
                self.source_file.set(values['source_file'])
                self.destination_ou.set(values['destination_ou'])
                self.domain.set(values['domain'])
                self.upn_suffix.set(values['upn_suffix'])
                self.result_file.set(values['result_file'])
                self.logfile.set(values['logfile'])
        except FileNotFoundError:
            pass

    def _browse_filename(self):
        file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=(("Excel workbooks", "*.xlsx"),))

        if file_path:
            self.source_file.set(file_path)

    def update_language(self, code):
        set_language(code)
        for widget in self.master.winfo_children():
            if widget.winfo_class() in ('Label', 'Button'):
                text = widget.cget('text')
                if text:
                    widget.config(text=_(text))
