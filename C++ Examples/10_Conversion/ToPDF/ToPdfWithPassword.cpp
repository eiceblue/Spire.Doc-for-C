#include "../pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToPdfWithPassword.pdf";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//create a parameter
	intrusive_ptr<ToPdfParameterList> toPdf = new ToPdfParameterList();

	//set the password
	std::wstring password = L"E-iceblue";
	toPdf->GetPdfSecurity()->Encrypt(L"password", password.c_str(), PdfPermissionsFlags::Default, PdfEncryptionKeySize::Key128Bit);
	//save doc file.
	document->SaveToFile(outputFile.c_str(), toPdf);
	document->Close();
}