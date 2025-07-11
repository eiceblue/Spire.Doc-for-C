#include "pch.h"
#include <fstream>
#include <locale>
#include <codecvt>

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/GetRowCellIndex.txt";
	std::wstring inputFile = DATAPATH"/ReplaceTextInTable.docx";

	//Load Word from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Get the first table in the section
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	std::wstring content;

	//Get table collections
	intrusive_ptr<TableCollection> collections = section->GetTables();

	//Get the table index
	int tableIndex = collections->IndexOf(table);

	//Get the index of the last table row
	intrusive_ptr<TableRow> row = table->GetLastRow();
	int rowIndex = row->GetRowIndex();

	//Get the index of the last table cell
	intrusive_ptr<TableCell> cell = Object::Dynamic_cast<TableCell>(row->GetLastChild());
	int cellIndex = cell->GetCellIndex();

	//Append these information into content
	content.append(L"Table index is " + std::to_wstring(tableIndex) + L"\r\n");
	content.append(L"Row index is " + std::to_wstring(rowIndex) + L"\r\n");
	content.append(L"Cell index is " + std::to_wstring(cellIndex) + L"\r\n");

	//Save to txt file
	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << content;
	write.close();

	doc->Close();
}
