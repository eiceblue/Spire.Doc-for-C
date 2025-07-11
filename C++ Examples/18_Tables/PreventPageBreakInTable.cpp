#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/PreventPageBreakInTable.docx";
	std::wstring inputFile = DATAPATH"/Template_Docx_5.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the table from Word document.
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(document->GetSections()->GetItemInSectionCollection(0)->GetTables()->GetItemInTableCollection(0));

	//Change the paragraph setting to keep them together.
	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(i);
		for (int j = 0; j < row->GetCells()->GetCount(); j++)
		{
			intrusive_ptr<TableCell> cell = row->GetCells()->GetItemInCellCollection(j);
			for (int k = 0; k < cell->GetParagraphs()->GetCount(); k++)
			{
				intrusive_ptr<Paragraph> p = cell->GetParagraphs()->GetItemInParagraphCollection(k);
				p->GetFormat()->SetKeepFollow(true);
			}
		}
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
