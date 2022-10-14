import os
import re
from collections import defaultdict
from typing import Generator, Union

PATH_TYPE = Union[str, bytes, os.PathLike]


class Answer:
    def __init__(self, file: PATH_TYPE = None):
        self.data = defaultdict(lambda: defaultdict(list))
        if file:
            self.load(file)

    def load(self, file: PATH_TYPE):
        with open(file, "r", encoding="utf-8") as f:
            lines = f.readlines()
        self.parse(lines)

    def parse(self, lines: list[str]):

        self.data.clear()

        for line in lines:

            line = line.rstrip("\n")

            # 空行・コメント行はスキップ
            if re.match(r"^\s*(#.+)?\s*$", line):
                continue

            values = line.replace(" ", "").split("\t")
            s = values[0]  # シート
            c = values[1]  # セル
            v = values[2]  # 解答例

            self.data[s][c].append(v)

    def sheets(self) -> Generator[str, None, None]:
        for s in self.data:
            yield s

    def cells(self, sheet: str) -> Generator[str, None, None]:
        for c in self.data[sheet]:
            yield c

    def match(self, sheet: str, cell: str, value: str) -> bool:

        # 数式のスペースを除去
        v = re.sub(r"\s+", "", value)

        # 解答データのリストと照合
        for a in self.data[sheet][cell]:
            # 解答データに一致するものが見つかった場合
            if re.fullmatch(a, v):
                return True

        # どの解答例にも当てはまらなかった場合
        # ただし，想定外の解答方法の可能性もある
        return False
