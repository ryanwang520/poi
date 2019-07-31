from .nodes import Box, BoxInstance
from .visitors.printer import print_visitor
from .visitors.writer import writer_visitor
from .writer import Writer


class Sheet:
    def __init__(self, root: Box):
        BoxInstance(root, 0, 0, None)
        self.root = root

    def write_to_bytesio(self) -> Writer:
        writer = Writer()
        visitor = writer_visitor(writer)
        self.root.accept(visitor)
        writer.close()
        return writer

    def write(self, filename):
        writer = self.write_to_bytesio()
        with open(filename, "wb") as f:
            f.write(writer.read())

    def print(self):
        self.root.accept(print_visitor)

    def to_bytes_io(self):
        writer = Writer()
        visitor = writer_visitor(writer)
        self.root.accept(visitor)
        writer.close()
        return writer.output
