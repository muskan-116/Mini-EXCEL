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
    gotoxy(89, 1);
    cout << "Use Arrow Keys to move." << endl;
    gotoxy(89, 2);
    cout << "Press 'k' to insert Value to CurrentActice." << endl;
    gotoxy(89, 3);
    cout << "Press 'i' to Add Row Above the CurrentActive." << endl;
    gotoxy(89, 4);
    cout << "Press 'h' to Add Row Below the CurrentActive." << endl;
    gotoxy(89, 5);
    cout << "Press 'y' to Add Column to the Right of the CurrentActive." << endl;
    gotoxy(89, 6);
    cout << "Press 'z' to Add Column to the Left of the CurrentActive." << endl;
    gotoxy(89, 7);
    cout << "Press 'l' to Insert Cell by Right Shift." << endl;
    gotoxy(89, 8);
    cout << "Press 'e' to Insert Cell by Down Shift." << endl;
    gotoxy(89, 9);
    cout << "Press 'g' to Delete Current Row." << endl;
    gotoxy(89, 10);
    cout << "Press 'f' to Clear Current Row." << endl;
    gotoxy(89, 11);
    cout << "Press 't' to Delete Current Column." << endl;
    gotoxy(89, 12);
    cout << "Press 'w' to Clear Current Column." << endl;
    gotoxy(89, 13);
    cout << "Press 'r' to Calculate Range Sum." << endl;
    gotoxy(89, 14);
    cout << "Press 'p' to Calculate Range Average." << endl;
    gotoxy(89, 15);
    cout << "Press 'x' to Calculate Range Count." << endl;
    gotoxy(89, 16);
    cout << "Press 's' to Calculate Range Minimum." << endl;
    gotoxy(89, 17);
    cout << "Press 'u' to Calculate Range Maximum." << endl;
    gotoxy(89, 18);
    cout << "Press 'v' to Copy Data." << endl;
    gotoxy(89, 19);
    cout << "Press 'b' to Cut Data." << endl;
    gotoxy(89, 20);
    cout << "Press 'q' to Paste Data." << endl;
    gotoxy(89, 21);
    cout << "Press 'n' to Load Data from File." << endl;
    gotoxy(89, 22);
    cout << "Press 'a' to Store Data in File." << endl;
    gotoxy(89, 23);
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



#include<iostream>
#include<conio.h>
#include<vector>
#include <iomanip>
#include<fstream>
#include<sstream>
#include<Windows.h>
#pragma once
using namespace std;
template <typename T>

class MiniEXCEL
{

	class Cell
	{
	
		T data;
		Cell* left;
		Cell* right;
		Cell* down;
		Cell* up;
		friend class MiniEXCEL;
	public:
		Cell(T val)
		{
			data = val;
			up = nullptr;
			down = nullptr;
			left = nullptr;
			right = nullptr;
		}
	};
	Cell* root;
	Cell* currentActive;
	
	int currentR;
	int currentC;

	int numRows;
	int numCols;
public:
	
	Cell* createGrid(int row, int column)
	{
		Cell* head = new Cell(1);
		Cell* current = head;
		for (int i = 0; i < row; ++i)
		{
			Cell* currentRow = current;
			for (int j = 0; j < column; ++j)
			{
				Cell* newCell = new Cell(1);
				currentRow->right = newCell;
				newCell->left = currentRow;
				currentRow = newCell;
				if (j > 0)
				{
					newCell->left = currentRow->left;
					currentRow->left->right = newCell;
				}

			}
			if (i > 0)
			{
				Cell* prevRow = current->up;
				prevRow->down = current;
				current->up = prevRow;
			}

			if (i < numRows - 1)
			{
				Cell* newRow = new Cell(1);
				current->down = newRow;
				newRow->up = current;
				current = newRow;
			}
		}
		return head;
	}

	MiniEXCEL(int rows, int cols)
	{
		numRows = rows;
		numCols = cols;
		currentR = 0;
		currentC = 0;
		currentActive = GetCell(0,0);
		root = createGrid(rows, cols);
		
	}
	class Iterator 
	{
	
		friend class MiniEXCEL;
	private:
		Cell* currentCell;
		int numRows;
		int numCols;
	public:
		Iterator(Cell* cell, int rows, int cols) 
		{
			currentCell = cell;
			numRows = rows;
			numCols = cols;
		}

		Iterator& operator++() {
			// Pre-increment operator (move right)
			if (currentCell->right) {
				currentCell = currentCell->right;
				numCols++;
			}
			return *this;
		}

		Iterator& operator--() {
			// Pre-decrement operator (move left)
			if (currentCell->left) {
				currentCell = currentCell->left;
				numCols--;
			}
			return *this;
		}

		Iterator operator++(int) {
			// Post-increment operator (move down)
			Iterator temp = *this;
			if (currentCell->down) {
				currentCell = currentCell->down;
				numRows++;
			}
			return temp;
		}

		Iterator operator--(int) {
			// Post-decrement operator (move up)
			Iterator temp = *this;
			if (currentCell->up) {
				currentCell = currentCell->up;
				numRows--;
			}
			return temp;
		}

		T& operator*() {
			// Dereference operator to access the data in the cell
			return currentCell->data;
		}

		bool operator==(const Iterator& other) const
		{
			return currentCell == other.currentCell;
		}

		bool operator!=(const Iterator& other) const
		{
			return !(*this == other);
		}
		Iterator& getCurrentCell()
		{
			return root;
		}
		
	};

	void printGrid() {
		Cell* currentRow = root;
		for (int i = 0; i < numRows; ++i) {
			
			Cell* currentCol = currentRow;

			for (int j = 0; j < numCols; ++j)
			{
				
				if (i == currentR && j == currentC) {
					cout << "{" << currentCol->data << "}";
					
				}
				else {
					cout << currentCol->data << "\t";
					
				}
				currentCol = currentCol->right;
				
			}
			
			cout << endl;
			
			currentRow = currentRow->down;
		}
	}

	void PrintCellBorder(int width)
	{
		for (int i = 0; i < width; ++i)
		{
			cout << "*^^^";
		}
		cout << "^" << endl;
	}

	void PrintDataInCell(const Cell* cell)
	{
		cout << setw(4) << left << cell->data;
	}

	void PrintGrid()
	{
		const int cellWidth = 4;

		Cell* currentRow = root;
		PrintCellBorder(numCols);

		for (int i = 0; i < numRows; ++i)
		{
			Cell* currentCol = currentRow;
			cout << "|";

			for (int j = 0; j < numCols; ++j)
			{
				if (i == currentR && j == currentC)
				{
					cout << "{";
					PrintDataInCell(currentCol);
					cout << "}";
				}
				else
				{
					PrintDataInCell(currentCol);
				}

				cout << "|";
				currentCol = currentCol->right;
			}

			cout << endl;
			PrintCellBorder(numCols);
			currentRow = currentRow->down;
		}
	}

	void PrintSheet()
	{
		PrintGrid();
	}
	void moveUp()
	{
		if (currentR > 0 )
		{
			
			currentActive = GetCell(currentR - 1, currentC);
			currentR--;
			currentActive = currentActive->left;
		}
	}

	void moveDown()
	{
		if (currentR < numRows - 1 )
		{
			
			currentActive = GetCell(currentR + 1, currentC);
			currentR++;
			
		}
	}

	void moveLeft()
	{
		if (currentC > 0 )
		{
			currentActive = GetCell(currentR, currentC - 1);
			currentC--;
			
		}
	}

	void moveRight()
	{
		if (currentC < numCols - 1 )
		{
			
			 currentActive = GetCell(currentR, currentC + 1);
			 currentC++;
			
		}
	}


	void InsertRowAbove() {
		if (currentR == 0)
		{  // Check if you are at the first row
			// Special case: Inserting a row above the first row
			Cell* newRow = createGrid(1, numCols);
			newRow->down = root;
			if (root) {
				root->up = newRow;
			}
			root = newRow;
			numRows++;
		}
		else {
			// Find the cell at the start of the previous row
			Cell* previousRow = root;
			for (int i = 0; i < currentR - 1; ++i) {
				previousRow = previousRow->down;
			}

			// Create a new row and set up and down links
			Cell* newRow = createGrid(1, numCols);
			newRow->down = previousRow->down;
			if (previousRow->down) {
				previousRow->down->up = newRow;
			}
			previousRow->down = newRow;
			newRow->up = previousRow;

			// If you were already at the last row, update the root pointer
			if (currentR == numRows - 1) {
				root = newRow;
			}

			numRows++;
		}
	}

	void InsertRowBelow() {
		if (currentR < numRows - 1) {  // Check if you are not at the last row

			// Find the cell at the start of the current row
			Cell* currentRow = root;
			for (int i = 0; i < currentR; ++i) {
				currentRow = currentRow->down;
			}

			// Find the cell at the start of the next row
			Cell* nextRow = currentRow->down;

			// Create a new row with cells containing default values (1 in this case)
			Cell* newRow = createGrid(1, numCols);

			// Update down links for the cells in the current row
			Cell* currentCell = currentRow;
			Cell* newCell = newRow;
			for (int j = 0; j < numCols; ++j) {
				newCell->down = currentCell->down;
				if (currentCell->down != nullptr) {
					currentCell->down->up = newCell;
				}
				currentCell->down = newCell;
				newCell->up = currentCell;
				currentCell = currentCell->right;
				newCell = newCell->right;
			}

			// Update down links for the rows above and below the new row
			newRow->down = nextRow;
			if (nextRow != nullptr) {
				nextRow->up = newRow;
			}
			currentRow->down = newRow;

			// Increment the total number of rows
			numRows++;
		}
	}
	void InsertValue(int val)
	{
		// Find the current cell
		Cell* currentRow = root;
		for (int i = 0; i < currentR; ++i) {
			currentRow = currentRow->down;
		}

		Cell* currentCell = currentRow;
		for (int j = 0; j < currentC; ++j) {
			currentCell = currentCell->right;
		}

		// Update the value in the current cell
		currentCell->data = val;
	}
	
		void ClearGrid() {
		// Start from the root and delete each cell and its columns
		Cell* currentRow = root;
		while (currentRow) {
			Cell* currentCell = currentRow;
			while (currentCell) {
				Cell* next = currentCell->right;
				delete currentCell;
				currentCell = next;
			}
			Cell* nextRow = currentRow->down;
			delete currentRow;
			currentRow = nextRow;
		}

		// Reset variables
		root = nullptr;
		currentR = 0;
		currentC = 0;
		numRows = 0;
		numCols = 0;
	}

	void DeleteRow() {
		if (numRows == 0) {
			// The grid is empty, nothing to delete
			return;
		}

		if (numRows == 1) {
			// Only one row is left, clear the grid and reset variables
			ClearGrid();
			return;
		}

		// Find the cell at the start of the current row
		Cell* currentRow = root;
		Cell* previousRow = nullptr;

		for (int i = 0; i < currentR; ++i) {
			previousRow = currentRow;
			currentRow = currentRow->down;
		}

		// Update up links for the cells in the current row
		Cell* currentCell = currentRow;
		for (int j = 0; j < numCols; ++j) {
			if (currentCell->up) {
				currentCell->up->down = currentCell->down;
			}
			if (currentCell->down) {
				currentCell->down->up = currentCell->up;
			}
			currentCell = currentCell->right;
		}

		// Update the root pointer if the deleted row is the first row
		if (currentR == 0) {
			root = root->down;
		}

		// Adjust variables
		if (previousRow) {
			previousRow->down = currentRow->down;
		}

		delete currentRow; // Free memory for the deleted row
		currentR = max(currentR - 1, 0);
		numRows--;
	}
	void ClearColumn()
	{
		if (currentActive)
		{
			// Find the top cell of the current column
			Cell* top = currentActive;
			while (top->up)
			{
				top = top->up;
			}

			// Clear the data in the column
			while (top)
			{
				top->data = T();  // Set to the default value for type T
				top = top->down;
			}
		}
	}

	void ClearRow() {
		// Find the cell at the start of the current row
		Cell* currentRow = root;
		for (int i = 0; i < currentR; ++i) {
			currentRow = currentRow->down;
		}

		// Find the first cell of the current row
		Cell* currentCell = currentRow;

		// Iterate through the cells in the current row and clear their data
		for (int j = 0; j < numCols; ++j) {
			currentCell->data = T();  // Clear the data (assign a default-constructed value)
			currentCell = currentCell->right;
		}
	}
	void deleteColumn()
	{
		if (currentActive && currentR >= 0 && currentR < numRows && currentC >= 0 && currentC < numCols)
		{
			// Find the current column
			Cell* currentCol = root;
			Cell* previousCol = nullptr;

			// Move to the current column
			for (int i = 0; i < currentC; ++i)
			{
				previousCol = currentCol;
				currentCol = currentCol->right;
			}

			// Get the column to the right of the current column
			Cell* nextCol = currentCol->right;

			// Update the links of the neighboring columns
			if (previousCol)
			{
				previousCol->right = nextCol;
			}
			else
			{
				root = nextCol; // If deleting the first column, update the head
			}

			if (nextCol)
			{
				nextCol->left = previousCol;
			}

			// Delete the cells in the current column
			Cell* currentCell = currentCol;
			Cell* nextCell;

			while (currentCell)
			{
				nextCell = currentCell->down;
				delete currentCell;
				currentCell = nextCell;
			}

			// Update the total number of columns
			--numCols;

			// Move the current active cell to the cell on the left if it's not in the first column
			if (currentC > 0)
			{
				currentActive = previousCol;
				--currentC;
			}
			else if (numCols > 0)
			{
				currentActive = root; // Move to the new first column if it exists
				--currentC; // Adjust currentY to columns - 1
			}
			else
			{
				currentActive = nullptr; // Set to nullptr if the grid becomes empty
			}
		}
	}
	void InsertColumnToRight()
	{
		if (currentActive)
		{
			// Move to the top of the current column
			Cell* top = currentActive;
			while (top->up)
			{
				top = top->up;
			}

			// Create a new cell for the top of the new column
			Cell* newColTop = new Cell(6);
			Cell* current = newColTop;

			// Iterate through the existing cells in the column
			while (top)
			{
				// Insert the new column to the right of each existing cell
				if (top->right == nullptr)
				{
					top->right = newColTop;
					newColTop->left = top;
				}
				else
				{
					// Connect the new column to the right of the existing cells
					Cell* rightCell = top->right;
					top->right = newColTop;
					newColTop->left = top;
					newColTop->right = rightCell;
					if (rightCell)  // Check if rightCell is not nullptr
					{
						rightCell->left = newColTop;
					}
				}

				// Move down to the next row
				top = top->down;

				if (top)
				{
					// Create a new cell for the next row in the new column
					Cell* nextRow = new Cell(6);
					current->down = nextRow;
					nextRow->up = current;
					current = nextRow;
				}
			}

			// Move to the next column
			newColTop = newColTop->right;

			numCols++;
			currentC++;
		}
	}
	//void InsertColumnToLeft()
	//{
	//	if (root == nullptr)
	//	{
	//		// If the grid is empty, create a new cell and set it as the root
	//		root = createGrid(0, 0);
	//	}
	//	else
	//	{
	//		// Create a new column at the beginning of each row
	//		Cell* newColumn = new Cell();
	//		newColumn->right = root;
	//		root->left = newColumn;
	//		newColumn->down = nullptr;

	//		// Iterate through each row and create a new cell
	//		while (root->down != nullptr)
	//		{
	//			Cell<T>* newRowCell = new Cell<T>();
	//			root = root->down;
	//			root->left = newRowCell;
	//			newRowCell->right = root;
	//			newRowCell->up = root->up->left;
	//			root->up->left->down = newRowCell;
	//		}

	//		// Move the pointer back to the first cell in the first column
	//		while (root->up != nullptr)
	//		{
	//			root = root->up;
	//		}
	//	}
	//}

	
	Iterator begin()
	{
		return Iterator(root ,numRows, numCols );
	}
	Iterator end()
	{
		return Iterator(nullptr, numRows, numCols);
	}
	public:
		void SUM(int startRow, int startCol, int endRow, int endCol) {
			// Calculate the sum using the GetRangeSum function
			int sum = GetRangeSum(startRow, startCol, endRow, endCol);

			// Set the sum value in the current active cell
			SetCellValue(sum);
		}

		void AVG(int startRow, int startCol, int endRow, int endCol) {
			// Calculate the average using the GetRangeAverage function
			int average = GetRangeAverage(startRow, startCol, endRow, endCol);

			// Set the average value in the current active cell
			SetCellValue(average);
		}

		void COUNT(int startRow, int startCol, int endRow, int endCol) {
			// Calculate the count using the range dimensions
			int count = (endRow - startRow + 1) * (endCol - startCol + 1);

			// Set the count value in the current active cell
			SetCellValue(count);
		}

		void MIN(int startRow, int startCol, int endRow, int endCol) {
			// Calculate the minimum using the GetRangeMin function
			int min = GetRangeMin(startRow, startCol, endRow, endCol);

			// Set the minimum value in the current active cell
			SetCellValue(min);
		}

		void MAX(int startRow, int startCol, int endRow, int endCol) {
			// Calculate the maximum using the GetRangeMax function
			int max = GetRangeMax(startRow, startCol, endRow, endCol);

			// Set the maximum value in the current active cell
			SetCellValue(max);
		}

		

		int GetRangeMin(int startRow, int startCol, int endRow, int endCol) {
			// Initialize the minimum to the maximum possible integer value
			int min = 10000;

			// Iterate through the range and update the minimum
			for (int i = startRow; i <= endRow; ++i) {
				for (int j = startCol; j <= endCol; ++j) {
					int value = GetValueAt(i, j);
					if (value < min)
					{
						min = value;
					}
				}
			}

			return min;
		}

		int GetRangeMax(int startRow, int startCol, int endRow, int endCol) {
			// Initialize the maximum to the minimum possible integer value
			int max = -100000;

			// Iterate through the range and update the maximum
			for (int i = startRow; i <= endRow; ++i) {
				for (int j = startCol; j <= endCol; ++j) {
					int value = GetValueAt(i, j);
					if  (value > max)
					{
						max = value;
					}
				}
			}

			return max;
		}	
		int GetRangeSum(int startRow, int startCol, int endRow, int endCol) {
		// Initialize the sum to zero
		int sum = 0;

		// Iterate through the range and accumulate the sum
		for (int i = startRow; i <= endRow; ++i) {
			for (int j = startCol; j <= endCol; ++j) {
				sum += GetValueAt(i, j);
			}
		}

		return sum;
	}

	int GetRangeAverage(int startRow, int startCol, int endRow, int endCol) {
		// Calculate the sum using the GetRangeSum function
		int sum = GetRangeSum(startRow, startCol, endRow, endCol);

		// Calculate the count using the range dimensions
		int count = (endRow - startRow + 1) * (endCol - startCol + 1);

		// Calculate the average (sum divided by count)
		if (count == 0) {
			// Handle the case where the count is zero to avoid division by zero
			return 0;  // Return a default value for type int
		}
		else {
			return sum / count;
		}
	}

	// Add the SetCellValue function to set the value in the current active cell
	void SetCellValue(int value) {
		// Find the current cell
		Cell* currentRow = root;
		for (int i = 0; i < currentR; ++i) {
			currentRow = currentRow->down;
		}

		Cell* currentCell = currentRow;
		for (int j = 0; j < currentC; ++j) {
			currentCell = currentCell->right;
		}

		// Update the value in the current cell
		currentCell->data = value;
	}

	// Add the GetValueAt function to get the value at a specific cell
	int GetValueAt(int row, int col) {
		// Find the cell at the specified row and column
		Cell* currentRow = root;
		for (int i = 0; i < row; ++i) {
			currentRow = currentRow->down;
		}

		Cell* currentCell = currentRow;
		for (int j = 0; j < col; ++j) {
			currentCell = currentCell->right;
		}

		// Return the value in the specified cell
		return currentCell->data;
	}
	
	vector<T> Copy(int startRow, int startCol, int endRow, int endCol) {
		// Validate the range indices to ensure they are within the grid's boundaries
		if (startRow < 0 || startCol < 0 || endRow >= numRows || endCol >= numCols) {
			// Handle invalid range indices, such as out-of-bounds indices
			throw std::out_of_range("Invalid range indices");
		}

		// Calculate the total number of elements to copy
		int elementsToCopy = (endRow - startRow + 1) * (endCol - startCol + 1);

		// Create a vector to store the copied data
		vector<T> copiedData(elementsToCopy);

		// Index to keep track of the position in the copiedData vector
		int index = 0;

		// Iterate through the range and copy the data to the vector
		for (int i = startRow; i <= endRow; ++i) {
			for (int j = startCol; j <= endCol; ++j) {
				copiedData[index++] = GetValueAt(i, j);
			}
		}

		return copiedData;
	}
	Cell* GetCell(int row, int col) const
	{
		// Validate the row and column indices
		if (row < 0 || row >= numRows || col < 0 || col >= numCols) {
			throw out_of_range("Invalid cell indices");
		}

		// Find the cell at the specified row and column
		Cell* currentRow = root;
		for (int i = 0; i < row; ++i) {
			currentRow = currentRow->down;
		}

		Cell* currentCol = currentRow;
		for (int j = 0; j < col; ++j) {
			currentCol = currentCol->right;
		}

		// Return a pointer to the specified cell
		return currentCol;
	}
	vector<int> Cut(int startRow, int startCol, int endRow, int endCol)
	{
		vector <int> cutted;
		for (int i = startRow; i <= endRow; i++)
		{
			for (int j = startCol; j <= endCol; j++)
			{
				Cell* current = GetCell(i, j);
				if (current)
				{
					cutted.push_back(current->data);
					current->data = 0;
				}
			}
		}
		return cutted;
	}
	void Paste(const vector<int>& data, int startRow, int startCol)
	{
		// Validate the startRow and startCol indices
		if (startRow < 0 || startRow >= numRows || startCol < 0 || startCol >= numCols)
		{
			throw out_of_range("Invalid start indices");
		}

		int dataIndex = 0;
		currentActive = GetCell(0,0);

		
		for (int i = startRow; i < numRows && dataIndex < data.size(); ++i)
		{
			
			for (int j = startCol; j < numCols && dataIndex < data.size(); ++j)
			{
				
				GetCell(i, j)->data = data[dataIndex++];
			}
			
		}

		// If there are remaining elements in the data vector, add more rows/columns
		while (dataIndex < data.size())
		{
			if (numRows <= startRow)
			{
				// Insert a new row if needed
				InsertRowBelow();
			}

			if (numCols <= startCol)
			{
				// Insert a new column if needed
				InsertColumnToRight();
			}

			// Update the value in the current cell with the copied data
			GetCell(startRow, startCol)->data = data[dataIndex++];

			// Move to the next cell
			++startCol;

			if (startCol >= numCols)
			{
				// Move to the next row if the end of the column is reached
				++startRow;
				startCol = 0;
			}
		}
	}
	void InsertCellByRightShift()
	{
		// Check if the current cell is at the rightmost column
		if (currentC < numCols - 1)
		{
			// Move to the right cell
			currentActive = currentActive->right;
			currentC++;

			// Create a new cell to insert to the left of the current cell
			Cell* newCell = new Cell(0);

			// Update the links to insert the new cell
			newCell->left = currentActive->left;
			newCell->right = currentActive;
			currentActive->left->right = newCell;
			currentActive->left = newCell;
		}
		else
		{
			// If the current cell is at the rightmost column, insert a new column to the right
			InsertColumnToRight();

			// Move to the first cell in the new column
			currentActive = GetCell(currentR, currentC);
		}
	}

	
	void LoadData(const string& filename)
	{
		ifstream file(filename , ios::in);
		if (file.is_open())
		{
			// Read numRows, numCols, currentR, currentC from the file
			file >> numRows >> numCols >> currentR >> currentC;

			// Create a new grid with the loaded dimensions
		
			root = createGrid(numRows, numCols);
			currentActive = GetCell(0, 0);
			// Iterate through each cell and read its data from the file
			for (int i = 0; i < numRows; ++i) {
				for (int j = 0; j < numCols; ++j) {
					file >> GetCell(i, j)->data;
				}
			}

			file.close();
		}
		else
		{
			cerr << "Error opening file for reading: " << filename << endl;
		}
	}
	void StoreData(const string& filename)
	{
		ofstream file(filename);
		if (file.is_open()) 
		{
			// Write numRows, numCols, currentR, currentC to the file
			file << numRows << " " << numCols << " " << currentR << " " << currentC << "\n";

			// Iterate through each cell and write its data to the file
			for (int i = 0; i < numRows; ++i) {
				for (int j = 0; j < numCols; ++j) {
					file << GetCell(i, j)->data << " ";
				}
				file << "\n";
			}

			file.close();
		}
		else 
		{
			cerr << "Error opening file for writing: " << filename << endl;
		}
	}
	
};
