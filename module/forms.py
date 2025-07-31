import gettext
import pickle
import tkinter as tk
from tkinter import messagebox, font, filedialog

from application import LANGUAGE_CODE, LAST_VALUES, LOCALE_DIR


def set_language(code):
    global LANGUAGE_CODE
    global _
    LANGUAGE_CODE = code
    lang = gettext.translation(
        'application', localedir=LOCALE_DIR, languages=[code], fallback=True)
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
        self.protocol = tk.StringVar(value='LDAP')

        # Load last values
        self._load_last_values()

        # Fonts
        label_font = font.Font(family='Times New Roman',
                               size=14, weight='bold')
        entry_font = font.Font(family='Times New Roman', size=14)

        # labels and entries
        widgets = [
            ("AD Server name:", self.ldap_server),
            ("Username:", self.username),
            ("Password:", self.password),
            ("Source filename:", self.source_file),
            ("Destination OU:", self.destination_ou),
            ("Domain:", self.domain),
            ("UPN-suffix:", self.upn_suffix),
            ("Results file name:", self.result_file),
            ("Log file name:", self.logfile)
        ]

        for row, (label_text, variable) in enumerate(widgets):
            label = tk.Label(master, text=label_text, font=label_font)
            label.grid(sticky='e', row=row, column=0, padx=(30, 5), pady=10)

            if label_text == "Password:":
                entry = tk.Entry(master, show='*', width=40,
                                 textvariable=variable, font=entry_font)
            else:
                entry = tk.Entry(master, width=40,
                                 textvariable=variable, font=entry_font)
            entry.grid(row=row, column=1, padx=(5, 30), pady=10, ipady=3)

        # Dynamic row calculation for buttons and dropdowns
        next_row = len(widgets)

        # buttons
        browse_button = tk.Button(master, text='\U0001F4C2', command=self._browse_filename,
                                  font=font.Font(family='Times New Roman', size=12))
        browse_button.grid(row=3, column=1, sticky='e', padx=(0, 30))

        button_start = tk.Button(
            master, text="Start import", command=self._on_start, font=label_font)
        button_start.grid(row=next_row, columnspan=2,
                          column=0, padx=50, pady=15, ipadx=129)

        button_cancel = tk.Button(
            master, text="Cancel", command=self._on_cancel, font=label_font)
        button_cancel.grid(row=next_row+1, columnspan=2,
                           column=0, padx=50, pady=15, ipadx=150)

        # Language Selection Dropdown
        self.language_var = tk.StringVar(value=LANGUAGE_CODE)
        self.language_var.trace_add(
            'write', lambda *args: self.update_language(self.language_var.get()))
        language_menu = tk.OptionMenu(
            self.master, self.language_var, "en_US", "ru_RU")
        language_menu.grid(row=next_row+2, column=0, padx=10, pady=5)

        # Protocol Selection Dropdown
        protocol_menu = tk.OptionMenu(
            self.master, self.protocol, 'LDAP', 'LDAPS')
        protocol_menu.grid(row=next_row+2, column=1,
                           sticky='e', padx=10, pady=5)

    def _on_start(self):
        """
        Requirement for confirmation of import
        Returns: None

        """
        # Show a confirmation dialog before proceeding
        confirm = messagebox.askyesno(
            _("Confirmation"), _(f"Do you want to start import?\n"))

        if confirm:
            self._save_last_values()
            self.master.destroy()  # Close the form window

    def _on_cancel(self):
        """
        Close window and exit program.
        Returns:

        """
        # Show a confirmation dialog before cancelling
        confirm = messagebox.askyesno(
            _("Confirmation"), _("Are you sure you want to cancel?"))

        if confirm:
            self.master.destroy()
            exit()  # Close the form window

    def _save_last_values(self):
        # Do not persist password for security reasons
        values = {
            'ldap_server': self.ldap_server.get(),
            'username': self.username.get(),
            'source_file': self.source_file.get(),
            'destination_ou': self.destination_ou.get(),
            'domain': self.domain.get(),
            'upn_suffix': self.upn_suffix.get(),
            'result_file': self.result_file.get(),
            'logfile': self.logfile.get(),
            # 'password': self.password.get(),  # Not saved for security
        }
        with open(LAST_VALUES, 'wb') as fp:
            pickle.dump(values, fp)

    def _load_last_values(self):
        try:
            with open(LAST_VALUES, 'rb') as fp:
                values = pickle.load(fp)
                self.ldap_server.set(values.get('ldap_server', ''))
                self.username.set(values.get('username', ''))
                self.source_file.set(values.get('source_file', ''))
                self.destination_ou.set(values.get('destination_ou', ''))
                self.domain.set(values.get('domain', ''))
                self.upn_suffix.set(values.get('upn_suffix', ''))
                self.result_file.set(values.get('result_file', ''))
                self.logfile.set(values.get('logfile', ''))
        except Exception:
            # If file is missing or corrupted, ignore and use defaults
            pass

    def _browse_filename(self):
        file_path = filedialog.askopenfilename(
            defaultextension=".xlsx", filetypes=(("Excel workbooks", "*.xlsx"),))

        if file_path:
            self.source_file.set(file_path)

    def update_language(self, code):
        set_language(code)
        for widget in self.master.winfo_children():
            if widget.winfo_class() in ('Label', 'Button'):
                text = widget.cget('text')
                if text:
                    widget.config(text=_(text))
