import os

import os
from datetime import datetime

def main():
    
    curr_dir = os.getcwd()
    
    print(curr_dir)
    print(os.listdir(curr_dir))
if __name__ == "__main__":
    main()
