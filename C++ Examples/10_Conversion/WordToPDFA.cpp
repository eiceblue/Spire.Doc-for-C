#include "pch.h"
using namespace Spire::Doc;


int main() {

	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"WordToPDFA.pdf";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Set the Conformance-level of the Pdf file to PDF_A1B.
	intrusive_ptr<ToPdfParameterList> toPdf = new ToPdfParameterList();
	toPdf->SetPdfConformanceLevel(PdfConformanceLevel::Pdf_A1B);

	//Save the file.
	document->SaveToFile(outputFile.c_str(), toPdf);
	document->Close();

}