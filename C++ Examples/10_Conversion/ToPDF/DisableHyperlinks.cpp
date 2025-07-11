#include "../pch.h"
using namespace Spire::Doc;


int main() {

	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_5.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DisableHyperlinks.pdf";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();


	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Create an instance of ToPdfParameterList.
	intrusive_ptr<ToPdfParameterList> pdf = new ToPdfParameterList();

	//Set DisableLink to true to remove the hyperlink effect for the result PDF page. 
	//Set DisableLink to false to preserve the hyperlink effect for the result PDF page.
	pdf->SetDisableLink(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), pdf);
	document->Close();
}
