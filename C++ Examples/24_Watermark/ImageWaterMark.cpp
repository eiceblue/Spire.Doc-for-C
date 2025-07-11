#include "pch.h"
using namespace Spire::Doc;

void InsertImageWatermark(intrusive_ptr<Document> document)
{
	intrusive_ptr<PictureWatermark> picture = new PictureWatermark();
	picture->SetPicture(DATAPATH"/ImageWatermark.png");
	picture->SetScaling(250);
	picture->SetIsWashout(false);
	document->SetWatermark(picture);
}

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ImageWaterMark.docx";

	//Open a Word document as template.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Insert the imgae watermark.
	InsertImageWatermark(document);
	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}




