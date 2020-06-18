#!/usr/bin/env python
import sys
import os
import pandas as pd
import re

from collections import OrderedDict
from typing import Generator, List

from difflib import unified_diff
from oletools.olevba3 import VBA_Parser
import colorama

from colorama import Fore, Style


def get_sheets(workbook: str, sep: str = u"\u2551", **kwargs):
    """
    Read the entire excel workbook and return a dict of stringified content

    >>> a = get_sheets("test01.xlsb")
    >>> len(a)
    9
    >>> "Resource" in a
    True
    >>> len(a["Resource"])
    193
    >>> len(a["Resource"][10].split("|"))
    37
    >>> "Queue" in a
    True
    >>> len(a["Queue"])
    31
    >>> len(a["Queue"][10].split("|"))
    11
    >>> b = get_sheets("xlsvba/helloworld.xlsm", sep = "  |  ")
    >>> len(b)
    2
    >>> "Sheet2" in b
    True
    >>> len(b["Sheet2"][0].split("  |  "))
    2
    """
    engine = "pyxlsb" if workbook.endswith(".xlsb") else None
    sheets = pd.read_excel(
        workbook,
        sheet_name=None,
        keep_default_na=False,
        header=None,
        engine=engine,
        **kwargs,
    )
    return {
        sheet_name: [
            sep.join([str(col) for col in row]) for i, row in df.iterrows()
        ]
        for sheet_name, df in sheets.items()
    }


def get_vba(workbook):
    vba_parser = VBA_Parser(workbook)
    vba_modules = (
        vba_parser.extract_all_macros()
        if vba_parser.detect_vba_macros()
        else []
    )

    modules = {}

    for _, _, _, content in vba_modules:
        decoded_content = content.decode("latin-1")
        lines = []
        if "\r\n" in decoded_content:
            lines = decoded_content.split("\r\n")
        else:
            lines = decoded_content.split("\n")
        if lines:
            name = lines[0].replace("Attribute VB_Name = ", "").strip('"')
            content = [
                line
                for line in lines[1:]
                if not (line.startswith("Attribute") and "VB_" in line)
            ]
            modules[name] = "\n".join(content)
    return modules


def colorize_diff_lines(diff_gen: Generator[str, None, None]) -> List[str]:
    return [
        (
            Fore.RED
            if line.startswith("-")
            else (
                Fore.GREEN
                if line.startswith("+")
                else (Fore.CYAN if line.startswith("@") else Style.RESET_ALL)
            )
        )
        + line.strip("\n")
        for line in list(diff_gen)
        if not line.startswith("---") and not line.startswith("+++")
    ]


if __name__ == "__main__":
    if len(sys.argv) != 8:
        print("Unexpected number of arguments")
        sys.exit(0)

    _, workbook_name, workbook_b, _, _, workbook_a, _, _ = sys.argv

    path_workbook_a = os.path.abspath(workbook_a)
    path_workbook_b = os.path.abspath(workbook_b)

    workbook_a_modules = (
        {} if workbook_a == os.devnull else get_vba(path_workbook_a)
    )
    workbook_b_modules = (
        {} if workbook_b == os.devnull else get_vba(path_workbook_b)
    )
    workbook_a_sheets = (
        {} if workbook_a == os.devnull else get_sheets(path_workbook_a)
    )
    workbook_b_sheets = (
        {} if workbook_b == os.devnull else get_sheets(path_workbook_b)
    )

    diffs = []
    for module_a, vba_a in workbook_a_modules.items():
        if module_a not in workbook_b_modules:
            diffs.append(
                {
                    "a": "--- /dev/null",
                    "b": "+++ b/" + workbook_name + "/VBA/" + module_a,
                    "diff": "\n".join(
                        colorize_diff_lines(
                            unified_diff(
                                [],
                                vba_a.split("\n"),
                            )
                        )
                    ),
                }
            )
        elif vba_a != workbook_b_modules[module_a]:
            diffs.append(
                {
                    "a": "--- a/" + workbook_name + "/VBA/" + module_a,
                    "b": "+++ b/" + workbook_name + "/VBA/" + module_a,
                    "diff": "\n".join(
                        colorize_diff_lines(
                            unified_diff(
                                workbook_b_modules[module_a].split("\n"),
                                vba_a.split("\n"),
                            )
                        )
                    ),
                }
            )

    for module_b, vba_b in workbook_b_modules.items():
        if module_b not in workbook_a_modules:
            diffs.append(
                {
                    "a": "--- a/" + workbook_name + "/VBA/" + module_b,
                    "b": "+++ /dev/null",
                    "diff": "\n".join(
                        colorize_diff_lines(
                            unified_diff(
                                vba_b.split("\n"),
                                [],
                            )
                        )
                    ),
                }
            )

    sheets = OrderedDict.fromkeys(
        [
            sheet
            for workbook in [workbook_a_sheets, workbook_b_sheets]
            for sheet in workbook.keys()
        ]
    )
    for sheet in sheets:
        label_b = "+++ " + (
            f"b/{workbook_name}/{sheet}"
            if sheet in workbook_a_sheets
            else "/dev/null"
        )
        label_a = "--- " + (
            f"a/{workbook_name}/{sheet}"
            if sheet in workbook_b_sheets
            else "/dev/null"
        )
        # to follow the convention of the original code above
        # b is a and a is b
        diff_ba = unified_diff(
            workbook_b_sheets.get(sheet, []), workbook_a_sheets.get(sheet, [])
        )
        diffs.append(
            {
                "a": label_a,
                "b": label_b,
                "diff": "\n".join(colorize_diff_lines(diff_ba)),
            }
        )

    colorama.init(strip=False)

    print(
        Style.BRIGHT
        + "diff --xltrail "
        + "a/"
        + workbook_name
        + " b/"
        + workbook_name
    )
    for diff in diffs:
        print(Style.BRIGHT + diff["a"])
        print(Style.BRIGHT + diff["b"])
        print(diff["diff"])
        print(Style.RESET_ALL)
