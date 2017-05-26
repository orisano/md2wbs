import xlwt
import mistune


class ExcelRenderer(mistune.Renderer):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self._current_line = 0
        self._workbook = xlwt.Workbook()
        self._sheet = self._workbook.add_sheet("md2xls")
        self._prev_level = 0
        self._current_id = [0] * 3

    def header(self, text, level, raw=None):
        self._current_id[level - 1] += 1
        for i in range(level, 3):
            self._current_id[i] = 0
        self._writeln(level, text)
        self._prev_level = level
        return ""

    def list_item(self, text):
        self._current_id[2] += 1
        self._write(self._prev_level + 1, text)
        self._writeln(self._prev_level + 2, ".".join(map(str, self._current_id)))
        return ""

    def _write(self, x, body):
        self._sheet.write(self._current_line, x - 1, body)

    def _writeln(self, x, body):
        self._sheet.write(self._current_line, x - 1, body)
        self._current_line += 1

    def save(self, fname):
        self._workbook.save(fname)


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("markdown")
    parser.add_argument("output")
    args = parser.parse_args()

    md = open(args.markdown).read()
    renderer = ExcelRenderer()
    mistune.markdown(md, renderer=renderer)
    renderer.save(args.output)


if __name__ == "__main__":
    main()
