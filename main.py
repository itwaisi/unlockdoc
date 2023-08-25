import os
from unlockdoc.unlockdoc import unlock


if __name__ == "__main__":
    
    passwords = os.path.join(os.getcwd(), '.temp', 'passwords.txt')
    
    protected_file = os.path.join(os.getcwd(), '.temp', 'protected.docx')
    
    unlock(protected_file, passwords)
