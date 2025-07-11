#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/WordToPdfEncrypt.pdf";
	std::wstring inputFile = DATAPATH"/Template_Docx_2.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Create an instance of ToPdfParameterList.
	intrusive_ptr<ToPdfParameterList> toPdf = new ToPdfParameterList();

	//Set the user password for the resulted PDF file.
	toPdf->GetPdfSecurity()->Encrypt(L"e-iceblue");

	//Save to file.
	document->SaveToFile(outputFile.c_str(), toPdf);
	document->Close();
}
