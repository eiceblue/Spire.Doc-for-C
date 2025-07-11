#include "../pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SpecifyEmbeddedFont.pdf";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	document->LoadFromFile(inputFile.c_str());
	//Specify embedded font
	intrusive_ptr<ToPdfParameterList> parms = new ToPdfParameterList();
	std::vector<std::wstring> part;
	part.push_back(L"PT Serif Caption");
	parms->SetEmbeddedFontNameList(part);
	document->SaveToFile(outputFile.c_str(), parms);
	document->Close();
}