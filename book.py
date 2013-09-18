class LibraryBook:

    def __init__(self, isbn, record, title, author, date, edition, url):
        self.isbn = isbn
        self.record = record
        self.title = title
        self.author = author
        self.date = date
        self.edition = edition
        self.url = url

    def __eq__(self, other):
        for sisbn in self.isbn:
            if sisbn != 'none':
                if sisbn in other.isbn:
                    if (self.title == other.title) or (self.author == other.author):
                        if self.edition == other.edition:
                            if self.date == other.date:
                                return True
        
        if self.record == other.record:
            if self.title == other.title:
                if self.author == other.author:
                    if self.date == other.date:
                        if self.url == other.url:
                            if self.edition == other.edition:
                                return True
        return False
