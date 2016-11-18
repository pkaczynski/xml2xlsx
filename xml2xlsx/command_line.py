# -*- coding: utf-8 -*-
import sys

import xml2xlsx


def main():
    if sys.platform == "win32":
        import os, msvcrt
        msvcrt.setmode(sys.stdout.fileno(), os.O_BINARY)

    sys.stdout.write(xml2xlsx.xml2xlsx(sys.stdin.read()))
