#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddPageBorders.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Define the border style.
	document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->GetBorders()->SetBorderType(BorderStyle::DotDash);

	//Define the border color.
	document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->GetBorders()->SetColor(Color::GetRed());

	//Set the line width.
	document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->GetBorders()->SetLineWidth(2);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}