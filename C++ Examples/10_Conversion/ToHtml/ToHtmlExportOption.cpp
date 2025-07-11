#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToHtmlTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToHtmlExportOption.html";

	//Open a Word document.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());
	//Set whether the css styles are embeded or not. 
	document->GetHtmlExportOptions()->SetCssStyleSheetFileName(L"sample.css");
	document->GetHtmlExportOptions()->SetCssStyleSheetType(CssStyleSheetType::External);

	//Set whether the images are embeded or not. 
	document->GetHtmlExportOptions()->SetImageEmbedded(false);
	document->GetHtmlExportOptions()->SetImagesPath(output_path.c_str());

	//Set the option whether to export form fields as plain text or not.
	document->GetHtmlExportOptions()->SetIsTextInputFormFieldAsText(true);

	//Save the document to a html file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Html);
	document->Close();
}
