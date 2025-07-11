#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/AddBarcodeImage.docx";
	std::wstring inputFile = DATAPATH"/SampleB_2.docx";

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	std::wstring imgPath = DATAPATH"/barcode.png";

	//Add barcode image
	intrusive_ptr<DocPicture> picture = document->GetSections()->GetItemInSectionCollection(0)->AddParagraph()->AppendPicture(imgPath.c_str());

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}