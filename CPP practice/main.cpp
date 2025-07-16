#include "Repo.h"
//using std::endl;
int main()
{
    const string bookData[] =
    {
        "NX_1, Forensics for Housewvies, Gen Val, 23.4, 45.99, 35, NONFICTION",
        "Fx0_F2, Useless Forensics for Housewvies, Gen Val, 23.4, 45.99, 35, FICTION"
    };

    const int numBooks = 2;
    Repo repo;

    for (int i = 0; i < numBooks; i++) repo.parse(bookData[i]);
    cout << "Displaying all books: " << std:: endl;
    repo.printAll();
    cout << std::endl;

    for (int i = 0, i < 3; i++)
    {
        cout << "Displaying by book type: " << bookTypeStrings[i] << std::endl;
        repo.printByBookType((BookType)i);
    }

    cout << "Displaying books with invalid IDs" << std:: endl;
    repo.printInvalidIDs();
    cout << std::endl;

    cout << "Displaying average prices: " << std::endl;
    repo.printAveragePrices();

    cout << "Removing book with ID N_W1:" << std::endl;
    repo.removeBookByID("N_W1");
    cout << std::endl;
    
    cout << "Removing book with ID N_W1:" << std::endl;
    repo.removeBookByID("N_W1");
    cout << std::endl;

    system("pause");
    return 0;
}