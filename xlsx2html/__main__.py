import sys
from . import xlsx2html

if __name__ == "__main__":
    if len(sys.argv[1:]) < 2:
        print("Usage: xlsx2html input.xlsx output.html")
        sys.exit()
    xlsx2html(sys.argv[1], sys.argv[2])
