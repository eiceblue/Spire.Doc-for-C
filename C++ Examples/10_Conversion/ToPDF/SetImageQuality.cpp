#include "../pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Doc_1.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetImageQuality.pdf";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str(), FileFormat::Doc);

	//Set the output image quality to be 40% of the original image. The default set of the output image quality is 80% of the original.
	document->SetJPEGQuality(40);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	document->Close();
}