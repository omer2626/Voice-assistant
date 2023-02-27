import pandas as pd
import speech_recognition as sr
import pyttsx3

# Load the Excel sheet into a pandas DataFrame
df = pd.read_excel(r"C:\\Users\\DELL\\Downloads\\syllabus1.xlsx")

# Set up a recognizer object
z = sr.Recognizer()

# Initialize the text-to-speech engine
engine = pyttsx3.init()


# Prompt the user to choose a class
#engine.say('Please say the name of the class you want to know the syllabus for:')
print("Please say the name of the class you want to know the syllabus for:")
with sr.Microphone() as source:
    audio = z.listen(source)
class_name = z.recognize_google(audio).lower()

# Filter the DataFrame to get only the rows for the chosen class
class_df = df[df['Class'] == class_name.capitalize()]
print(class_name.capitalize())

# Read out the available subjects for the chosen class
subjects = ', '.join(class_df['Subject'].unique())
engine.say(f"The subjects for {class_name.capitalize()} are: {subjects}")
# engine.say('maths,social')
engine.runAndWait()

# Prompt the user to choose a subject
print("Please say the name of the subject you want to know the syllabus for:")
with sr.Microphone() as source:
    audio = z.listen(source)
subject_name = z.recognize_google(audio).lower()

print(subject_name.capitalize())

# Filter the DataFrame further to get only the rows for the chosen subject
subject_df = class_df[class_df['Subject'] == subject_name.capitalize()]
print(subject_df)

# Read out the topics for the chosen subject
topics = subject_df['Topic'].tolist()
engine.say(f"The topics for {subject_name.capitalize()} are ")
for topic in topics:
    engine.say(f'{topic}')
    print(topic)
# engine.say('World countries, india, our culture')
engine.runAndWait()