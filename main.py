from PyQt5 import QtWidgets
import sys
import pandas as pd
import os

class EmotionCollector(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        # Text Entry
        self.text_label = QtWidgets.QLabel("Enter Malayalam Text:")
        self.text_entry = QtWidgets.QTextEdit()

        # Emotion Dropdown
        self.emotion_label = QtWidgets.QLabel("Select Emotion:")
        self.emotion_combo = QtWidgets.QComboBox()
        self.emotion_combo.addItems(["Happy", "Excitement", "Sad", "Sarcasm", "Humour", "Anger", "Love", "Surprise", "Abusive", "Fear"])

        # Submit Button
        self.submit_button = QtWidgets.QPushButton("Submit")
        self.submit_button.clicked.connect(self.submit_data)

        # Layout
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.text_label)
        layout.addWidget(self.text_entry)
        layout.addWidget(self.emotion_label)
        layout.addWidget(self.emotion_combo)
        layout.addWidget(self.submit_button)

        self.setLayout(layout)
        self.setWindowTitle("Malayalam Text Emotion Dataset Collector")

    def submit_data(self):
        text = self.text_entry.toPlainText().strip()
        emotion = self.emotion_combo.currentText()

        if text == "":
            QtWidgets.QMessageBox.critical(self, "Input Error", "Please enter some text.")
            return

        if emotion == "":
            QtWidgets.QMessageBox.critical(self, "Input Error", "Please select an emotion.")
            return

        # Prepare data
        data = pd.DataFrame({"Text": [text], "Emotion": [emotion]})

        try:
            if os.path.exists("emotion_dataset.xlsx"):
                df = pd.read_excel("emotion_dataset.xlsx")
                df = pd.concat([df, data], ignore_index=True)
            else:
                df = data

            df.to_excel("emotion_dataset.xlsx", index=False)

            # Clear input fields
            self.text_entry.clear()
            self.emotion_combo.setCurrentIndex(0)

            QtWidgets.QMessageBox.information(self, "Success", "Data submitted successfully!")

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"An error occurred: {e}")

app = QtWidgets.QApplication(sys.argv)
window = EmotionCollector()
window.show()
sys.exit(app.exec_())
