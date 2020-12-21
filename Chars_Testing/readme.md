# Chars_Testing

This macro is based on the testing suite developed by Andrew Mason, Jack Dunn
and byi649.

It attempts to complement the work done by Sam Gilmour in identifying 
special characters which cause errors in OpenSolver.

His work can be found here:
https://github.com/OpenSolver/OpenSolver/issues/268

It is hoped that future developers of OpenSolver will be able to use the
macro to more easily identify scenarios where special characters cause errors,
rather than manually trying to find these.

# How to add more tests:

- Open the macro 'OpenSolver SheetNameCharTester.xlsm'. 
- Go to the rightmost worksheet in the workbook. e.g.: 'BadName&'
- Make a copy of it and select (move to end) in the menu which appears.
- Rename it according to which character you want to test e.g.: '%BadName'.
- Update the description in the worksheet accordingly.
- Repeat this process for the special character in the middle and at the end 
of the sheet name (e.g.: for Bad%Name and BadName%).
- Open the VBA developer. In this, open the module 'TestDispatcher'. 
Add to the cases already there as fits the tests you are adding.
- Save this code. Go back to the worksheet and insert hyperlinks on the 
'Results' page to the corresponding worksheets.
