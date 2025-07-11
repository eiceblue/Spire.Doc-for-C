#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TextBoxTable.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReadTableFromTextBox.txt";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first textbox
	intrusive_ptr<TextBox> textbox = doc->GetTextBoxes()->GetItem(0);

	//Get the first table in the textbox
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(textbox->GetBody()->GetTables()->GetItemInTableCollection(0));

	wstring stringBuilder;

	//Loop through the paragraphs of the table cells and extract them to a .txt file
	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(i);
		for (int j = 0; j < row->GetCells()->GetCount(); j++)
		{
			intrusive_ptr<TableCell> cell = row->GetCells()->GetItemInCellCollection(j);
			for (int k = 0; k < cell->GetParagraphs()->GetCount(); k++)
			{
				intrusive_ptr<Paragraph> paragraph = cell->GetParagraphs()->GetItemInParagraphCollection(k);
				stringBuilder.append(paragraph->GetText());
				stringBuilder.append(L"\t");
			}
		}
		stringBuilder.append(L"\n");
	}

	//Save to TXT file and launch it
	wofstream foo(outputFile);
	foo << stringBuilder.c_str();
	foo.close();
	doc->Close();
}
