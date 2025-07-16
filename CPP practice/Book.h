#pragma once
#include <iostream>
#include <iomanip>
#include "BookType.h"
using std::string;
using std::cout;

class Book
{
public: 
    const static int price ArraySize = 3;
private:
    string bookID;
    string title; 
    string author;
    double prices[priceArraySize];
    BookType bookType;
public:
    Book();
    Book(string bookID, string title, string author, double prices[], BookType bookType);
    ~Book();

    string getID();
    string getTitle()
    string getAuthor();
    double* getPrices();
    BookType getBookType();

    void setID(string ID);
    void setTitle(string title);
    void setAuthor(string Author);
    void setPrices(double prices[]);
    void setBookType(BookType bt);

    static void printHeader();

    void print();
}