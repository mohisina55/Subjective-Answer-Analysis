import openpyxl
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import utils

class Det_Report(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.master.title("Evaluation Full Report")
        self.master.geometry("700x650+361+100")
        self.master.configure(bg="white")

        # Destroy previous widgets before adding new content
        for widget in self.master.winfo_children():
            widget.destroy()

        main_frame = tk.Frame(self.master, bg="white")
        main_frame.pack(fill=tk.BOTH, expand=True)

        my_canvas = tk.Canvas(main_frame, bg="white")
        my_scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=my_canvas.yview)

        my_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        my_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        my_canvas.configure(yscrollcommand=my_scrollbar.set)
        my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))

        new_scroll = tk.Frame(my_canvas, bg="white")
        my_canvas.create_window((0, 0), window=new_scroll, anchor="nw")

        # Calculate scores
        similarity_score = sum(utils.all_similarity_scores) / len(utils.Qtext)
        grammar_score = sum(utils.all_grammar_scores) / len(utils.Qtext)
        spelling_score = sum(utils.all_spelling_scores) / len(utils.Qtext)
        keyword_score = sum(utils.all_keyword_scores) / len(utils.Qtext)

        # Save scores to utils for reference
        utils.similarity_score = similarity_score
        utils.grammar_score = grammar_score
        utils.spell_score = spelling_score
        utils.keyword_score = keyword_score

        # Display Scores
        tk.Label(new_scroll, text=f"Your Total Marks: {utils.total}", font='Arial 10 bold', bg="white").grid(row=1, column=1, padx=10, pady=5, sticky="w")
        tk.Label(new_scroll, text=f"Similarity Factor: {similarity_score:.2%}", font='Arial 10', bg="white").grid(row=2, column=1, padx=10, pady=5, sticky="w")
        tk.Label(new_scroll, text=f"Grammar Accuracy: {grammar_score:.2%}", font='Arial 10', bg="white").grid(row=3, column=1, padx=10, pady=5, sticky="w")
        tk.Label(new_scroll, text=f"Keyword Accuracy: {keyword_score:.2%}", font='Arial 10', bg="white").grid(row=4, column=1, padx=10, pady=5, sticky="w")
        tk.Label(new_scroll, text=f"Spelling Accuracy: {spelling_score:.2%}", font='Arial 10', bg="white").grid(row=5, column=1, padx=10, pady=5, sticky="w")

        # Display Question-wise Scores
        for i in range(len(utils.Qtext)):
            row_offset = 6 + (i * 5)
            tk.Label(new_scroll, text=f"Q{i+1} Score: {utils.all_final_score[i] * 10}", font='Arial 10 bold', bg="white").grid(row=row_offset, column=1, padx=10, pady=5, sticky="w")
            tk.Label(new_scroll, text=f"  Similarity: {utils.all_similarity_scores[i]:.2%}", font='Arial 10', bg="white").grid(row=row_offset+1, column=1, padx=10, pady=5, sticky="w")
            tk.Label(new_scroll, text=f"  Grammar: {utils.all_grammar_scores[i]:.2%}", font='Arial 10', bg="white").grid(row=row_offset+2, column=1, padx=10, pady=5, sticky="w")
            tk.Label(new_scroll, text=f"  Spelling: {utils.all_spelling_scores[i]:.2%}", font='Arial 10', bg="white").grid(row=row_offset+3, column=1, padx=10, pady=5, sticky="w")
            tk.Label(new_scroll, text=f"  Keywords: {utils.all_keyword_scores[i]:.2%}", font='Arial 10', bg="white").grid(row=row_offset+4, column=1, padx=10, pady=5, sticky="w")

        # Close Button
        tk.Button(new_scroll, text="Back", width=15, command=self.go_back).grid(row=row_offset+5, column=1, padx=10, pady=20)

    def go_back(self):
        """Go back to the previous report window"""
        from Report import Report
        Report(self.master)
