#!/bin/env python3

"""
Exampleapplication Helen
version 0
(c) 2021 T.Wisotzki
"""

import sys
import os.path
import tkinter as tk
from docx import Document


# global variable - defines where text-files are stored
TEXTS_PATH = "texts/"


class DocumentCreator(tk.Tk):
    """Main GUI"""

    def __init__(self, optionnames):
        super().__init__()
        self.title("Document creator")
        self.geometry("300x400")

        self.options = []
        self.choosen_options = []

        # instructions at top
        label = tk.Label(self, text="Select textblocks to include:")
        label.pack()

        # create checkboxes for every file
        for opname in optionnames:
            new_option = tk.Checkbutton(self, text=opname)
            new_option.bind("<Button-1>", self.add_option)
            self.options.append(new_option)
        for option in self.options:
            option.pack()

        # generation Button
        self.generate_button = tk.Button(
            self, text="generate document", command=self.generate_word
        )
        self.generate_button.pack()

        self.exit_button = tk.Button(self, text="exit", command=self.exit_app)
        self.exit_button.pack()

    def add_option(self, event):
        """if text-part is de-/selected it is added/removed"""
        task = event.widget.cget("text")
        if task in self.choosen_options:
            self.choosen_options.remove(task)
        else:
            self.choosen_options.append(task)

    @staticmethod
    def exit_app():
        """Exit"""
        sys.exit(0)

    def generate_word(self):
        """generation of combined file as word"""
        document = Document()
        document.add_heading("Example Document template")
        for option in self.choosen_options:
            with open(
                os.path.join(TEXTS_PATH, option + ".txt"), encoding="utf-8"
            ) as stream:
                paragraph = document.add_paragraph("")
                paragraph.add_run(stream.read())
        document.save("final.docx")


def get_options():
    """get all text-blocks in specified directory"""
    if not os.path.isdir(TEXTS_PATH):
        print("Error: path to text extracts not valid", file=sys.stderr)
        sys.exit(1)
    return [os.path.splitext(path)[0] for path in os.listdir(TEXTS_PATH)]


if __name__ == "__main__":
    documentcreator = DocumentCreator(get_options())
    documentcreator.mainloop()
