import tkinter as tk #Tkinter library for building the GUI
from tkinter import filedialog, messagebox #file dialog and messagebox for interaction
from FullAudit import FullAuditProgram
from WagerAudit import WagerAuditProgram
from JurisdictionGameVersionAudit import JurisdictionGameVersionAuditProgram
from GameVersionAudit import GameVersionAuditProgram


class AuditToolSelector:
    def __init__(self, master):
        self.master = master
        master.title("Audit Comparison Tool") #title
        master.configure(bg="#2b2b2b") #set window background color to white

        self.main_widgets() #function for UI components
        self.main_window() #function for screen function

        master.geometry("700x400")
        master.minsize(700, 400)

        master.protocol("WM_DELETE_WINDOW", self.close_window) #X button will confirm if user wants to close

    def main_window(self):
        #Get the screen's full width/height
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()

        #Defines the desired window dimensions
        window_width = 700
        window_height = 400

        #Calculate the top-left corner position to center the window
        position_top = int(screen_height / 2 - window_height / 2)
        position_left = int(screen_width / 2 - window_width / 2)

        #Update the window's geometry to apply size and position
        self.master.geometry(f'{screen_width}x{window_height}+{position_left}+{position_top}')

    def close_window(self): #Function for cancel confirmation
        confirm = messagebox.askyesno(
            "Exit Audit Comparison Tool",
            "Are you sure you want to close the Audit Comparison Tool?"
        )
        if confirm:
            self.master.destroy() #To close this window only
        else:
            messagebox.showinfo(
                "Canceled!",
                "Close canceled."
            )

    def main_widgets(self):
        #Main content frame for all buttons/labels
        content_frame = tk.Frame(self.master, bg="#2b2b2b", height=100)
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)

        #Welcome display text and label
        welcome_text = "\nAudit Comparison Tool\nSelect Audit Type\n"
        self.welcome_label = tk.Label(content_frame, text=welcome_text, font=("TkDefaultFont", 15, "bold"), fg='white', bg='#2b2b2b')
        self.welcome_label.pack(pady=10)

        #group container for buttons
        group_container = tk.Frame(content_frame, bg="#2b2b2b")
        group_container.pack()

        #Button style dictionary for all buttons
        button_style = {
            "bg": "#6e6e6e",
            "fg": "white",
            "activebackground": "#505050",
            "activeforeground": "white",
            "borderwidth": 1,
            "highlightthickness": 0,
            "font": ("TkDefaultFont", 10, "bold")
        }

        #Button style dictionary for exit button
        exit_button_style = {
            "borderwidth": 1,
            "highlightthickness": 0,
            "font": ("TkDefaultFont", 10, "bold")
        }

        #Wager and Game Version Audit button
        self.fullAudit_button = tk.Button(group_container, text="Wager & Game/Math Version Audit", width=35, command=lambda: FullAuditProgram(master=self.master), **button_style)
        self.fullAudit_button.pack(pady=(10))
        self.button_hover_effect(self.fullAudit_button)

        #Wager Audit button
        self.wagerAudit_button = tk.Button(group_container, text="Wager Audit", width=35, command=lambda: WagerAuditProgram(master=self.master), **button_style)
        self.wagerAudit_button.pack(pady=(10))
        self.button_hover_effect(self.wagerAudit_button)

        #Jurisdiction Game Version Audit button
        self.jurisdictionGameVersionAudit_button = tk.Button(group_container, text="Jurisdiction Game Version Audit", width=35, command=lambda: JurisdictionGameVersionAuditProgram(master=self.master), **button_style)
        self.jurisdictionGameVersionAudit_button.pack(pady=(10))
        self.button_hover_effect(self.jurisdictionGameVersionAudit_button)

        #Game Version Audit button
        self.gameVersionAudit_button = tk.Button(group_container, text="Game Version Audit", width=35, command=lambda: GameVersionAuditProgram(master=self.master), **button_style)
        self.gameVersionAudit_button.pack(pady=(10))
        self.button_hover_effect(self.gameVersionAudit_button)

        #Exit button
        self.exit_button = tk.Button(group_container, text="EXIT", width=20, command=self.close_window, bg = "#FF6F6F", fg = 'white', **exit_button_style)
        self.exit_button.pack(pady=(10))
        self.button_hover_effect(self.exit_button, normal_bg="#FF6F6F")

    #Adds a hover effect to buttons
    def button_hover_effect(self, button, hover_bg="#5a5a5a", normal_bg="#6e6e6e"):
        button.config(bg=normal_bg)

        def on_enter(e):
            if normal_bg == "#FF6F6F":
                button.config(bg="#cc0000")
            else:
                button.config(bg=hover_bg)

        def on_leave(e):
            button.config(bg=normal_bg)

        #Bind hover effect to the button
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

def main():
    root = tk.Tk()
    app = AuditToolSelector(root)
    root.mainloop()

if __name__ == "__main__":
    main()
