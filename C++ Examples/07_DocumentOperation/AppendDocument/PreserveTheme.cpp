#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Theme.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"PreserveTheme.docx";

	//Load the source document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Create a new Word document
	intrusive_ptr<Document> newWord = new Document();
	//Clone default style, theme, compatibility from the source file to the destination document
	doc->CloneDefaultStyleTo(newWord);
	doc->CloneThemesTo(newWord);
	doc->CloneCompatibilityTo(newWord);

	//Add the cloned section to destination document
	newWord->GetSections()->Add(doc->GetSections()->GetItemInSectionCollection(0)->CloneSection());

	//Save and launch document
	newWord->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	newWord->Close();
}