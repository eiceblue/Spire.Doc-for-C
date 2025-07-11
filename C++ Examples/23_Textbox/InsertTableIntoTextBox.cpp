#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertTableIntoTextBox.docx";

	//Create a new document
	intrusive_ptr<Document> doc = new Document();

	//Add a section
	intrusive_ptr<Section> section = doc->AddSection();

	//Add a paragraph to the section
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Add a textbox to the paragraph
	intrusive_ptr<TextBox> textbox = paragraph->AppendTextBox(300, 100);

	//Set the position of the textbox
	textbox->GetFormat()->SetHorizontalOrigin(HorizontalOrigin::Page);
	textbox->GetFormat()->SetHorizontalPosition(140);
	textbox->GetFormat()->SetVerticalOrigin(VerticalOrigin::Page);
	textbox->GetFormat()->SetVerticalPosition(50);

	//Add text to the textbox
	intrusive_ptr<Paragraph> textboxParagraph = textbox->GetBody()->AddParagraph();
	intrusive_ptr<TextRange> textboxRange = textboxParagraph->AppendText(L"Table 1");
	textboxRange->GetCharacterFormat()->SetFontName(L"Arial");

	//Insert table to the textbox
	intrusive_ptr<Table> table = textbox->GetBody()->AddTable(true);

	//Specify the number of rows and columns of the table
	table->ResetCells(4, 4);

	std::vector<std::vector<LPCWSTR_S>> data =
	{
		{L"Name", L"Age", L"Gender", L"ID"},
		{L"John", L"28", L"Male", L"0023"},
		{L"Steve", L"30", L"Male", L"0024"},
		{L"Lucy", L"26", L"female", L"0025"}
	};

	//Add data to the table 
	for (int i = 0; i < 4; i++)
	{
		for (int j = 0; j < 4; j++)
		{
			intrusive_ptr<TextRange> tableRange = table->GetRows()->GetItemInRowCollection(i)->GetCells()->GetItemInCellCollection(j)->AddParagraph()->AppendText(data[i][j]);
			tableRange->GetCharacterFormat()->SetFontName(L"Arial");
		}
	}

	//Apply style to the table
	table->ApplyStyle(DefaultTableStyle::TableColorful2);


	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
