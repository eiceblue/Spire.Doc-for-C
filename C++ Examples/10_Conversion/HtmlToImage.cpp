#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_HtmlFile1.html";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"HtmlToImage.png";


	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str(), FileFormat::Html, XHTMLValidationType::None);

	//Save to image in the default format of png.

	//intrusive_ptr<Image> image = document->SaveToImages(0, ImageType::Bitmap);
	//image->Save(outputFile.c_str(), FREE_IMAGE_FORMAT::FIF_PNG);
	intrusive_ptr<Stream> imageStream = document->SaveImageToStreams(0, ImageType::Bitmap);
	//Obtain image data in the default format of png,you can use it to convert other image format
	std::vector<byte> imageData = imageStream->ToArray();
	//TestUtil::WriteBytesToImage(imageData, outputFile,FREE_IMAGE_FORMAT::FIF_PNG);
	std::ofstream outFile(outputFile);
	if (outFile.is_open())
	{
		outFile.write(reinterpret_cast<const char*>(imageData.data()), imageData.size());
		outFile.close();
	}
	document->Close();
	imageStream->Dispose();

}