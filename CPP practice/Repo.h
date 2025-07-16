#pragma once
#include "Book.h"
class Repo
{
private:
    int lastIndex = -1;
    const static int numBook = 5;
    Book* bookRepoArray[numBooks];

public:
    void parse(string row);
    void add(string bID,
        string bTitle,
        stringbAuthor,
        double bprice1,
        double bprice2,
        double bprice3,
        BookType bt);
    void printAll();
    void printByBookType(BookType bt);
    void printInvalidIDs();
    void printAveragePrices();
    void removeBookById(string BookID);
    ~Repo();
};