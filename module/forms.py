import gettext
import pickle
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog

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
        self.ldap_server = ttk.StringVar()
        self.username = ttk.StringVar()
        self.password = ttk.StringVar()
        self.source_file = ttk.StringVar()
        self.destination_ou = ttk.StringVar()
        self.domain = ttk.StringVar()
        self.upn_suffix = ttk.StringVar()
        self.result_file = ttk.StringVar()
        self.logfile = ttk.StringVar()
        self.protocol = ttk.StringVar(value='LDAP')
        self.language_var = ttk.StringVar(value=LANGUAGE_CODE)
        self.theme_var = ttk.StringVar(value='darkly')

        # Load last values
        self._load_last_values()

        # Store widget references for language updates
        self.widgets_to_update = {}

        # Create top toolbar frame
        toolbar_frame = ttk.Frame(master, padding=(10, 5))
        toolbar_frame.grid(row=0, column=0, sticky=NSEW, padx=10, pady=(5, 0))

        # Configure grid weights
        master.columnconfigure(0, weight=1)
        master.rowconfigure(1, weight=1)

        # Top toolbar with dropdowns
        # Protocol dropdown
        protocol_label = ttk.Label(toolbar_frame, text="Protocol:", font=(
            "Times New Roman", 10))
        protocol_label.pack(side='left', padx=(30, 5))
        self.widgets_to_update['protocol_label'] = protocol_label

        protocol_menu = ttk.OptionMenu(
            toolbar_frame, self.protocol, self.protocol.get(), 'LDAP', 'LDAPS', style='info.TMenubutton')
        protocol_menu.pack(side='left', padx=(0, 20))

        # Language dropdown
        language_label = ttk.Label(toolbar_frame, text="Language:", font=(
            "Times New Roman", 10))
        language_label.pack(side='left', padx=(30, 5))
        self.widgets_to_update['language_label'] = language_label

        self.language_var.trace_add(
            'write', lambda *args: self.update_language(self.language_var.get()))
        language_menu = ttk.OptionMenu(
            toolbar_frame, self.language_var, self.language_var.get(), "en_US", "ru_RU", style='info.TMenubutton')
        language_menu.pack(side='left', padx=(0, 20))

        # Theme dropdown
        theme_label = ttk.Label(toolbar_frame, text="Theme:", font=(
            "Times New Roman", 10))
        theme_label.pack(side='left', padx=(30, 5))
        self.widgets_to_update['theme_label'] = theme_label

        self.theme_var.trace_add(
            'write', lambda *args: self._apply_theme())
        theme_menu = ttk.OptionMenu(
            toolbar_frame, self.theme_var, self.theme_var.get(
            ), 'darkly', 'cosmo', 'flatly', 'journal', 'litera',
            'lumen', 'minty', 'pulse', 'sandstone', 'simplex', 'solar', 'superhero', 'united',
            'yeti', style='info.TMenubutton')
        theme_menu.pack(side='left', padx=(0, 20))

        # Create main frame with padding
        main_frame = ttk.Frame(master, padding=20)
        main_frame.grid(row=1, column=0, sticky=(N, W, E, S), padx=10, pady=10)

        # Configure main frame grid weights
        main_frame.columnconfigure(1, weight=1)

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
            label = ttk.Label(main_frame, text=label_text,
                              font=("Times New Roman", 12, "bold"))
            label.grid(sticky='e', row=row, column=0, padx=(0, 10), pady=5)
            self.widgets_to_update[f'label_{row}'] = label

            if label_text == "Password:":
                entry = ttk.Entry(main_frame, show='*', width=40,
                                  textvariable=variable, font=("Times New Roman", 12))
            else:
                entry = ttk.Entry(main_frame, width=40,
                                  textvariable=variable, font=("Times New Roman", 12))
            entry.grid(row=row, column=1, padx=(0, 0), pady=5, sticky='ew')

        # Dynamic row calculation for buttons
        next_row = len(widgets)

        # Browse button
        browse_button = ttk.Button(main_frame, text='📁', command=self._browse_filename,
                                   style='secondary.TButton', width=3)
        browse_button.grid(row=3, column=2, padx=(5, 0), pady=5)

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=next_row, column=0, columnspan=3, pady=20)

        button_start = ttk.Button(
            button_frame, text="Start import", command=self._on_start,
            style='success.TButton', width=20)
        button_start.pack(pady=5)
        self.widgets_to_update['button_start'] = button_start

        button_cancel = ttk.Button(
            button_frame, text="Cancel", command=self._on_cancel,
            style='danger.TButton', width=20)
        button_cancel.pack(pady=5)
        self.widgets_to_update['button_cancel'] = button_cancel
        self.update_language(self.language_var.get())

    def _apply_theme(self):
        """Apply the selected theme to the window"""
        try:
            new_theme = self.theme_var.get()
            self.master.style.theme_use(new_theme)
        except Exception as e:
            print(f"Error applying theme: {e}")

    def _on_start(self):
        """
        Requirement for confirmation of import
        Returns: None

        """
        # Show a confirmation dialog before proceeding
        confirm = Messagebox.show_question(
            _("Confirmation"),
            _("Do you want to start import?"),
        )

        if confirm == "Yes":
            self._save_last_values()
            self.master.destroy()  # Close the form window

    def _on_cancel(self):
        """
        Close window and exit program.
        Returns:

        """
        # Show a confirmation dialog before cancelling
        confirm = Messagebox.show_question(
            _("Confirmation"),
            _("Are you sure you want to cancel?"),
        )

        if confirm == "Yes":
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
            'protocol': self.protocol.get(),
            'theme': self.theme_var.get(),
            'language': self.language_var.get(),
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
                self.protocol.set(values.get('protocol', ''))
                self.theme_var.set(values.get('theme', ''))
                self.language_var.set(values.get('language', ''))
        except Exception:
            # If file is missing or corrupted, ignore and use defaults
            pass

    def _browse_filename(self):
        file_path = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=(("Excel workbooks", "*.xlsx"),),
            parent=self.master
        )

        if file_path:
            self.source_file.set(file_path)

    def update_language(self, code):
        """Update the interface language"""
        set_language(code)

        # Update window title
        self.master.title(_("Mass input of users."))

        # Update stored widget texts
        translations = {
            'protocol_label': _("Protocol:"),
            'language_label': _("Language:"),
            'theme_label': _("Theme:"),
            'theme_button': _("Apply Theme"),
            'button_start': _("Start import"),
            'button_cancel': _("Cancel"),
            'label_0': _("AD Server name:"),
            'label_1': _("Username:"),
            'label_2': _("Password:"),
            'label_3': _("Source filename:"),
            'label_4': _("Destination OU:"),
            'label_5': _("Domain:"),
            'label_6': _("UPN-suffix:"),
            'label_7': _("Results file name:"),
            'label_8': _("Log file name:"),
        }

        # Update each widget
        for widget_name, widget in self.widgets_to_update.items():
            if widget_name in translations:
                widget.config(text=translations[widget_name])
