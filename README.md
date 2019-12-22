---

## How to run the application

1. You'll see the clone button under the **Source** heading. Click that button.
2. Open the solution **SimpleSpreadsheet.sln** with Visual Studio 2017.
3. Click the **Run** button.

---

## How to run the tests

1. Click the **Test** menu at the top of the **Visual Studio** tabs.
2. Click **Run** and **all tests** and there will be a test window appears.

---

## Design considerations

1. SimpleSpreadsheet.BLL (The project including business logic)
2. SimpleSpreadsheet.Common (The project including common modules and exceptions)
3. SimpleSpreadsheet (Console Application)
4. SimpleSpreadsheet.BLL.UnitTests (The project of Unit Tests)
6. SimpleSpreadsheet.DataAccess (Not really use this project as we don't have database access)

---

#Commands
Command 			Description
C w h           	Should create a new canvas of width w and height h.
L x1 y1 x2 y2   	Should create a new line from (x1,y1) to (x2,y2). Currently only
					horizontal or vertical lines are supported. Horizontal and vertical lines
					will be drawn using the 'x' character.
R x1 y1 x2 y2   	Should create a new rectangle, whose upper left corner is (x1,y1) and
					lower right corner is (x2,y2). Horizontal and vertical lines will be drawn
					using the 'x' character.
B x y c         	Should fill the entire area connected to (x,y) with "colour" c. The
					behavior of this is the same as that of the "bucket fill" tool in paint
					programs.
N x1 y1 v1			Should insert a number in specified cell (x1,y1)
S x1 y1 x2 y2 x3 y3 Should perform sum on top of all cells from x1 y1 to x2 y2 and store the result in x3 y3
Q               	Should quit the program.

