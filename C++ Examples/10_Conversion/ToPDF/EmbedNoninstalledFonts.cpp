#include "../pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"EmbedNoninstalledFonts.pdf";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Embed the non-installed fonts.
	intrusive_ptr<ToPdfParameterList> parms = new ToPdfParameterList();
	std::vector<intrusive_ptr<PrivateFontPath>> fonts;
	
	intrusive_ptr<PrivateFontPath> tempVar = new PrivateFontPath(L"PT Serif Caption", (input_path + L"PT_Serif-Caption-Web-Regular.ttf").c_str());
	fonts.push_back(tempVar);
	parms->SetPrivateFontPaths(fonts);

	//Save doc file to pdf.
	document->SaveToFile(outputFile.c_str(), parms);
	document->Close();
}