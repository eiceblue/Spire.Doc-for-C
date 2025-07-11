#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TextBoxTable.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DeleteTableFromTextBox.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first textbox
	intrusive_ptr<TextBox> textbox = doc->GetTextBoxes()->GetItem(0);

	//Remove the first table from the textbox
	textbox->GetBody()->GetTables()->RemoveAt(0);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
