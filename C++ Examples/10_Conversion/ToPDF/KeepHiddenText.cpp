#include "../pch.h"

using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_5.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"KeepHiddenText.pdf";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//When convert to PDF file, set the property IsHidden as true.
	intrusive_ptr<ToPdfParameterList> pdf = new ToPdfParameterList();
	pdf->SetIsHidden(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), pdf);
	document->Close();
}