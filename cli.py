# this file defines the entry point for pyinstaller for creating an *.exe package from the python project
# # to run pyinstaller: activate env, then run: pyinstaller cli.py --name bom_compare --onefile

from calculate_difference_bom import main

if __name__ == '__main__':
    main()