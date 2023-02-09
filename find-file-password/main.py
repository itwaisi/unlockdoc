import os
from win32com import client



def find_password():

    # PASSWORD PROTECTED WORD DOCUMENT
    file_pass = os.getcwd() + '\\' + 'document-test.docx'


    # PASSWORD FILE
    password_file = os.getcwd() + '/' + 'passwords-test.txt'


    # CORRECT PASSWORD FOR TESTING
    word_password = 'testpassword'


    # CREATE HEADLESS INSTANCE OF MICROSOFT WORD
    word = client.Dispatch("Word.Application")
    word.Visible = False


    # GET PASSWORDS FROM FILE/DB AS LIST
    passwords = []
    with open(password_file, 'r', encoding='UTF-8', errors='ignore') as file:
        passwords = [line.rstrip('\n') for line in file]


    # ADD WORKING PASSWORD TO LIST IF IT DOESN'T EXIST
    passwords.append(word_password)


    # TRY EACH PASSWORD
    for i, password in enumerate(passwords):
        if len(password) > 0:
            try:
                word.Documents.Open(file_pass, False, True, None, password)
                print(f'Password {i+1} - Password Length: {len(password)} - Correct Password: {password}')
                break
            except:
                print(f'Password {i+1} - Password Length: {len(password)} - Incorrect Password: {password}')

        if len(password) == 0:
            print(f'Password {i+1} - Password Length: {len(password)} - Incorrect Password: {password}')
    
    
    # CLOSE HEADLESS INSTANCE OF MICROSOFT WORD
    word.Quit()
