if __name__ == '__main__':
    import sys
    sys.path.insert(1, "../")

import os
from data2doc.data2doc import merge, data2doc


class TestCode:
    d2d = data2doc()

    def getTestFileNameXlsx(self):
        return f"{os.path.dirname(__file__)}/../test-data/test01/test.xlsx"

    def test_TestFileNameXlsx(self):
        assert os.path.isfile(self.getTestFileNameXlsx()) is True

    def getTestFileNameDocx(self):
        return f"{os.path.dirname(__file__)}/../test-data/test01/test.docx"

    def test_TestFileNameDocx(self):
        assert os.path.isfile(self.getTestFileNameDocx()) is True

    def getTestFileNameResult(self, fileName="result.docx"):
        tryName = f"{os.path.dirname(__file__)}/../test-data/test-result/{fileName}"
        if os.path.isfile(tryName):
            os.remove(tryName)
        return tryName

    def test_load_xlsx(self):
        assert self.d2d.loadXlsxFile(self.getTestFileNameXlsx()) is True

    def test_load_docx(self):
        assert self.d2d.loadDocxFile(self.getTestFileNameDocx()) is True

    def test_d2d(self):
        self.d2d.loadXlsxFile(self.getTestFileNameXlsx())
        self.d2d.loadDocxFile(self.getTestFileNameDocx())
        assert self.d2d.data2doc() is True
        assert len(self.d2d.docxResultBinary) > 0
        

    def test_merge(self):
        assert merge(self.getTestFileNameDocx(),
                     self.getTestFileNameXlsx(),
                     self.getTestFileNameResult()) is True


if __name__ == '__main__':
    Foo = TestCode()
    [getattr(Foo, meth)() for meth in dir(Foo) if callable(getattr(Foo, meth)) and meth.startswith("test")]
