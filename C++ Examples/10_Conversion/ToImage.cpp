#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToImage.png";

	//Create word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Stream> imageStream = document->SaveImageToStreams(0, ImageType::Bitmap);
	//Obtain image data in the default format of png,you can use it to convert other image format
	std::vector<byte> imgBytes = imageStream->ToArray();
	std::ofstream outFile(outputFile, std::ios::binary);
	if (outFile.is_open())
	{
		outFile.write(reinterpret_cast<const char*>(imgBytes.data()), imgBytes.size());
		outFile.close();
	}
	document->Close();
	imageStream->Dispose();


}