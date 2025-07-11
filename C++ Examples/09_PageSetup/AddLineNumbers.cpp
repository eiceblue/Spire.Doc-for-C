#include "pch.h"
using namespace Spire::Doc;


int main() {

	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddLineNumbers.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();


	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Set the start value of the line numbers.
	document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->SetLineNumberingStartValue(1);

	//Set the interval between displayed numbers.
	document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->SetLineNumberingStep(6);

	//Set the distance between line numbers and text.
	document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->SetLineNumberingDistanceFromText(40.0f);

	//Set the numbering mode of line numbers. There are four choices: None, Continuous, RestartPage and RestartSection.
	document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->SetLineNumberingRestartMode(LineNumberingRestartMode::Continuous);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}