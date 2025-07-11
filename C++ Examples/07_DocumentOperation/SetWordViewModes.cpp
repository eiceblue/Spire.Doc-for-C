#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring outputFile = output_path + L"SetWordViewModes.docx";

	//Create Word document.
	intrusive_ptr<Document> document =  new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Set Word view modes.
	document->GetViewSetup()->SetDocumentViewType(DocumentViewType::WebLayout);
	document->GetViewSetup()->SetZoomPercent(150);
	document->GetViewSetup()->SetZoomType(ZoomType::None);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
