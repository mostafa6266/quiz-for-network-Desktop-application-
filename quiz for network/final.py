import random
import tkinter as tk
import openpyxl
from Data_fore_ccna import data_for_ccna
from Data_for_ccnp import data_for_ccnp
class QuizApp:
    def __init__(self, master):
        self.master = master
        master.title("Quiz App")
        self.is_quiz_finished = False  # Variable to track if the quiz is finished
        self.user_answers = []  # List to store user's answers

        # Create the widgets
        self.created_label = tk.Label(master, text="Created by Eng:Moustafa Mohamed Mahmoud ", font=("Arial", 15), anchor="e")
        self.created_label.pack(side=tk.TOP, anchor="e", padx=10, pady=5)

        self.name_label = tk.Label(master, text="Enter Your Name:", font=("Arial", 14))
        self.name_label.pack(pady=5)
        self.name_entry = tk.Entry(master, font=("Arial", 14))
        self.name_entry.pack(pady=5)

        self.title_label = tk.Label(master, text="Welcome to the Quiz App!", font=("Arial", 20))
        self.start_button = tk.Button(master, text="Start Quiz", command=self.start_quiz, font=("Arial", 16))
        self.quit_button = tk.Button(master, text="Quit", command=self.quit_quiz, font=("Arial", 16), state=tk.DISABLED)  # Disabled by default

        # Layout the widgets
        self.title_label.pack()
        self.start_button.pack(pady=10)
        self.quit_button.pack(pady=10)

        # Disable the close button
        master.protocol("WM_DELETE_WINDOW", self.disable_close)


    def disable_close(self):
        if not self.is_quiz_finished:
            return  # Do nothing if quiz is not finished
        self.master.quit()

    def start_quiz(self):
        self.name_label.pack_forget()  # Hide the name label and entry field
        self.name_entry.pack_forget()

        # Destroy the start button and title label
        self.start_button.destroy()
        self.title_label.destroy()

        # Randomly selecting 50 questions for CCNA, 25 questions each for Helpdisk and MCSA
        # Determine the sample size for CCNA and CCNP questions
        ccna_keys = list(data_for_ccna.keys())
        ccnp_keys = list(data_for_ccnp.keys())

        # Set the desired sample sizes
        ccna_sample_size = min(50, len(ccna_keys))
        ccnp_sample_size = min(50, len(ccnp_keys))

        # Randomly select questions for CCNA and CCNP sections
        self.questions_for_ccna = random.sample(ccna_keys, ccna_sample_size)
        self.questions_for_helpdisk = random.sample(ccnp_keys, ccnp_sample_size)

        # Combine all the questions
        all_questions = self.questions_for_helpdisk + self.questions_for_ccna

        # Shuffle the combined questions
        random.shuffle(all_questions)


        # Printing the questions and keeping track of the score
        self.score = 0
        self.question_index = 0
        self.question_label = tk.Label(self.master, text="", font=("Arial", 18))
        self.question_label.pack(pady=10)
        self.choice_labels = []
        for i in range(4):
            choice_label = tk.Label(self.master, text="", font=("Arial", 14))
            self.choice_labels.append(choice_label)
            choice_label.pack(pady=5)
        self.answer_entry = tk.Entry(self.master, font=("Arial", 14))
        self.answer_entry.pack(pady=5)
        self.submit_button = tk.Button(self.master, text="Submit Answer", command=self.check_answer, font=("Arial", 14))
        self.submit_button.pack(pady=5)

        self.clear_screen()

    def clear_screen(self):
        # Clear the screen
        self.question_label.config(text="")
        self.answer_entry.delete(0, tk.END)
        for choice_label in self.choice_labels:
            choice_label.config(text="")

        # Get the next question
        if self.question_index < len(self.questions_for_ccna):
            question = self.questions_for_ccna[self.question_index]
            choices = data_for_ccna[question]
            self.question_label.config(text=question)
            for i in range(4):
                self.choice_labels[i].config(text=choices[i])

        elif self.question_index < len(self.questions_for_ccna) + len(self.questions_for_helpdisk):
            index = self.question_index - len(self.questions_for_ccna)
            question = self.questions_for_helpdisk[index]
            choices = data_for_ccnp[question]
            self.question_label.config(text=question)
            for i in range(4):
                self.choice_labels[i].config(text=choices[i])

   
        else:
            self.finish_quiz()

        self.question_index += 1

    def check_answer(self):
        # Check if the answer field is empty
        if not self.answer_entry.get():
            return  # Do nothing if the answer field is empty

        user_answer = self.answer_entry.get()
        self.user_answers.append(user_answer)  # Store user's answer

        # Check the answer and update the score
        user_answer = self.answer_entry.get()

        # Check if the current question is from the CCNA section
        if self.question_index <= len(self.questions_for_ccna):
            correct_answer = data_for_ccna[self.questions_for_ccna[self.question_index - 1]][-1]["Answer"]
            if user_answer.lower() == correct_answer.lower():
                self.score += 1  # Increment score by 1 if the answer is correct

        # Check if the current question is from the Helpdisk section
        elif self.question_index <= len(self.questions_for_ccna) + len(self.questions_for_helpdisk):
            index = self.question_index - len(self.questions_for_ccna) - 1
            correct_answer = data_for_ccnp[self.questions_for_helpdisk[index]][-1]["Answer"]
            if user_answer.lower() == correct_answer.lower():
                self.score += 1

        self.clear_screen()

    def save_result(self):
    # Get the name entered by the user
        name = self.name_entry.get()

        # Create a new Excel workbook
        workbook = openpyxl.Workbook()

        # Create a sheet for the main results
        sheet = workbook.active
        sheet.title = "Quiz Results"

        # Set the column headers for the main results sheet
        sheet['A1'] = "Question"
        sheet['B1'] = "Correct Answer"
        sheet['C1'] = "User's Answer"

        # Write the questions, correct answers, and user's answers to the main results sheet
        for i, question in enumerate(self.questions_for_ccna):
            correct_answer = data_for_ccna[question][-1]["Answer"]
            user_answer = self.user_answers[i] if i < len(self.user_answers) else ""

            sheet[f'A{i + 2}'] = question
            sheet[f'B{i + 2}'] = correct_answer
            sheet[f'C{i + 2}'] = user_answer

        for i, question in enumerate(self.questions_for_helpdisk, start=len(self.questions_for_ccna)):
            index = i - len(self.questions_for_ccna)
            correct_answer = data_for_ccnp[question][-1]["Answer"]
            user_answer = self.user_answers[i] if i < len(self.user_answers) else ""

            sheet[f'A{i + 2}'] = question
            sheet[f'B{i + 2}'] = correct_answer
            sheet[f'C{i + 2}'] = user_answer

        # Create a new sheet for final results
        final_sheet = workbook.create_sheet("Final Results")

        # Set the column headers for the final results sheet
        final_sheet['A1'] = "Name"
        final_sheet['B1'] = "Final Score"

        # Write the name and final score to the final results sheet
        final_sheet['A2'] = name
        final_sheet['B2'] = self.score

        # Save the workbook with a filename based on the name and current date/time
        filename = f"{name}_quiz_result.xlsx"
        workbook.save(filename)

   


    def finish_quiz(self):
        # Call the save_result method
        self.save_result()
        # Destroy the question widgets
        self.question_label.destroy()
        self.answer_entry.destroy()
        self.submit_button.destroy()
        for choice_label in self.choice_labels:
            choice_label.destroy()

        # Display the final score
        score_label = tk.Label(self.master, text="Your score: {}".format(self.score))
        score_label.pack()

        # Enable the close button
        self.quit_button.config(state=tk.NORMAL)

        # Set the is_quiz_finished variable to True
        self.is_quiz_finished = True


    def quit_quiz(self):
        self.master.quit()


# Create the main window
root = tk.Tk()

# Create the quiz app
app = QuizApp(root)

# Start the main event loop
root.mainloop()


