#include "Book.h"
Book::Book()
{
    this->bookId = "";
    this->title= "";
    this->author= "";
    for (int i = 0; i < price ArraySize; i++) this->prices[i] = 0;
    this->bookType = BookType::UNDECIDED;
}

Book::Book(string bookID, string title, str author, double prices[], BookType bookType)
{
    this->bookId = bookID;
    this->title= title;
    this->author= author;
    for (int i = 0; i < price ArraySize; i++) this->prices[i] = prices[i];
    this->bookType = bookType;   
}

Book::~Book() {}

string Book::getID() { return this->bookID; }
string Book::getTitle() { return this->title; }
string Book::getAuthor() { return this->author; }
double* Book:: getPrices() { return this -> prices; }
BookType Book::getBookType() {return this -> bookType; }

void Book::setID(string ID) { this->bookID = ID; }
void Book::setTitle(string title) { this->title = title;}
void Book::setAuthor(string author) { this->author = author;}
void Book::setPrices(double prices[])
{
    for (int i = 0; i < priceArraySize; i++) this->prices[i] = prices[i];
}
void Book::setBookType(BookType bt) { this->bookType = bt; }

void Book::printHeader()
{
    cout << "Output format: ID|Title|Author|Prices|Type\n";
}

void Book::print()
{
    cout << this->getID() << '\t';
    cout << this->getTitle() << '\t';
    cout << this->getAuthor() << '\t';
    cout << this->getPrices()[0] << '\t';
    cout << this->getPrices()[1] << '\t';
    cout << this->getPrices()[2] << '\t';
    cout << bookTypeStrings[this->getBookType()] << '\n';
}

