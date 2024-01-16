#include <iostream>
#include<fstream>
#include<windows.h>
#include<conio.h>
using namespace std;
#include"Header.h"

void gotoxy(int x, int y)
{
    COORD coordinates;
    coordinates.X = x;
    coordinates.Y = y;
    SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
}
void Menu()
{
    gotoxy(90, 1);
    cout << "Use Arrow Keys to move." << endl;
    gotoxy(90, 2);
    cout << "Press 'k' to insert Value to CurrentActice." << endl;
    gotoxy(90, 3);
    cout << "Press 'i' to Add Row Above the CurrentActive." << endl;
    gotoxy(90, 4);
    cout << "Press 'h' to Add Row Below the CurrentActive." << endl;
    gotoxy(90, 5);
    cout << "Press 'y' to Add Column to the Right of the CurrentActive." << endl;
    gotoxy(90, 6);
    cout << "Press 'z' to Add Column to the Left of the CurrentActive." << endl;
    gotoxy(90, 7);
    cout << "Press 'l' to Insert Cell by Right Shift." << endl;
    gotoxy(90, 8);
    cout << "Press 'e' to Insert Cell by Down Shift." << endl;
    gotoxy(90, 9);
    cout << "Press 'g' to Delete Current Row." << endl;
    gotoxy(90, 10);
    cout << "Press 'f' to Clear Current Row." << endl;
    gotoxy(90, 11);
    cout << "Press 't' to Delete Current Column." << endl;
    gotoxy(90, 12);
    cout << "Press 'w' to Clear Current Column." << endl;
    gotoxy(90, 13);
    cout << "Press 'r' to Calculate Range Sum." << endl;
    gotoxy(90, 14);
    cout << "Press 'p' to Calculate Range Average." << endl;
    gotoxy(90, 15);
    cout << "Press 'x' to Calculate Range Count." << endl;
    gotoxy(90, 16);
    cout << "Press 's' to Calculate Range Minimum." << endl;
    gotoxy(90, 17);
    cout << "Press 'u' to Calculate Range Maximum." << endl;
    gotoxy(90, 18);
    cout << "Press 'v' to Copy Data." << endl;
    gotoxy(90, 19);
    cout << "Press 'b' to Cut Data." << endl;
    gotoxy(90, 20);
    cout << "Press 'q' to Paste Data." << endl;
    gotoxy(90, 21);
    cout << "Press 'n' to Load Data from File." << endl;
    gotoxy(90, 22);
    cout << "Press 'a' to Store Data in File." << endl;
    gotoxy(90, 23);
    cout << "Press 'Esc.' to End." << endl;
}


int main()
{
    system("Color  03");
    cout << " __  __    ___    _  _     ___              ___   __  __    ___     ___     _    " << endl;
    cout << "|  \\/  |  |_ _|  | \\| |   |_ _|     o O O  | __|  \\ \\/ /   / __|   | __|   | | " << endl;
    cout << "| |\\/| |   | |   | .` |    | |     o       | _|    >  <   | (__    | _|    | |_" << endl;
    cout << "|_|__|_|  |___|  |_|\\_|   |___|   TS__[O]  |___|  /_/\\_\\   \\___|   |___|   |____| " << endl;
    cout << "|''''' | _| '''''|_|'''''  | _ | '''''| {======|_|''''' | _ | '''''|_|''''' | _ | '''''|_|''''' |  " << endl;
    cout << "`-0-0-''`-0 - 0 - ''`-0-0-''`-0-0-'.//o--000''`-0 - 0 - ''`-0-0-''`-0-0-''`-0 - 0 - '`-0-0-''" << endl;
    cout << endl;
    cout << "press any key to continue...." << endl;
    char h = _getch();
    system("cls");
    MiniEXCEL<int> excel(5, 5);
    int key;
    string path = "D:\\THIRD SEMESTER\\DSA\\PROJECT DSA\\ExcelProject\\grid.txt";
    system("Color 06");
    
    vector<int> copyData;
    excel.PrintSheet();
    Menu();
    
    
     while (true)
     {
         Menu();
        if (_kbhit()) 
        {
            
            key = _getch();
            switch (key)
            {
            case 'a':
                excel.StoreData(path);
                cout << "Data store successfully...." << endl;
                break;
            case 'n':
                excel.LoadData(path);
                cout << "Data loaded successdully";
                break;
            case 72:
                excel.moveUp();
                break;
            case 80:
                excel.moveDown();
                break;
            case 75:
                excel.moveLeft();
                break;
            case 77:
                excel.moveRight();
                break;
            case 'i':
                excel.InsertRowAbove();
                break;
            case 'h':
                excel.InsertRowBelow();
                break;
            case 'k':
                int value;
                cout << "Enter value:";
                cin >> value;
                excel.InsertValue(value);
                break;
            case 'y':
                excel.InsertColumnToRight();
                break;
             /*case 'z':
                 excel.InsertColumnToLeft();
                 break;*/
            case 'g':
                excel.DeleteRow();
                break;
            case 't':
                excel.deleteColumn();
                break;
            case 'f':
                excel.ClearRow();
                break;
            case 'w':
                excel.ClearColumn();
                break;
            
            case 'r':
                cout << "SUM " << endl;
                excel.SUM(1,1,2,2);
                break;
            case 'p':
                cout << "AVERAGE " << endl;
                excel.AVG(1, 1, 2, 2);
                break;
            case 'x':
                cout << "COUNT " << endl;
                excel.COUNT(1, 1, 2, 2);
                break;
            case 'u':
                cout << "MAX " << endl;
                excel.MAX(1, 1, 2, 2);
                break;
            case 's':
                cout << "MIN " << endl;
                excel.MIN(1, 1, 2, 2);
                break;
            case 'v':
                copyData = excel.Copy(1,1,2,2) ;
                for (int data : copyData)
                {
                    cout << data << "\t";
                }
                cout << endl;
                Sleep(200);
                break;
            case 'b':
                copyData = excel.Cut(1, 1, 2, 2);
                for (int data : copyData)
                {
                    cout << data << "\t";
                }
                cout << endl;
                Sleep(200);
                break;
            case 'q':
                copyData = excel.Copy(1, 1, 2, 2);
                excel.Paste(copyData, 4,4 );
                break;
            case 'l':
                excel.InsertCellByRightShift();
                break;
                    
            case 27: // ESC key to exit
                return 0;
            }
            system("cls");// Clear the console (Windows)
            excel.PrintSheet();
        }
     }
     
    return 0;

}



