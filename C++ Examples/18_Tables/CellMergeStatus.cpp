#include "pch.h"
#include <locale>
#include <codecvt>
#include <fstream>
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CellMergeStatus.txt";
	std::wstring inputFile = DATAPATH"/CellMergeStatus.docx";
	

	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Get the first table in the section
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	std::wstring stringBuidler;
	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		intrusive_ptr<TableRow> tableRow = table->GetRows()->GetItemInRowCollection(i);
		for (int j = 0; j < tableRow->GetCells()->GetCount(); j++)
		{
			intrusive_ptr<TableCell> tableCell = tableRow->GetCells()->GetItemInCellCollection(j);
			CellMerge verticalMerge = tableCell->GetCellFormat()->GetVerticalMerge();
			short horizontalMerge = tableCell->GetGridSpan();

			if (verticalMerge == CellMerge::None && horizontalMerge == 1)
			{
				stringBuidler.append(L"Row " + std::to_wstring(i) + L", cell " + std::to_wstring(j) + L": ");
				stringBuidler.append(L"This cell isn't merged.\n");
			}
			else
			{
				stringBuidler.append(L"Row " + std::to_wstring(i) + L", cell " + std::to_wstring(j) + L": ");
				stringBuidler.append(L"This cell is merged.\n");
			}
		}
		stringBuidler.append(L"\n");
	}

	//Save document
	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << stringBuidler;
	write.close();
	
	doc->Close();
}
