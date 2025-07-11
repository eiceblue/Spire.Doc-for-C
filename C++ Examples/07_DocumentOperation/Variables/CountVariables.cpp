#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_6.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CountVariables.txt";;
	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the number of variables in the document.
	int number = document->GetVariables()->GetCount();

	wstring content;
	content.append(L"The number of variables is: " + std::to_wstring(number));

	//Save to file.
	wofstream out;
	out.open(outputFile);
	out.flush();
	out << content.c_str();
	out.close();
	document->Close();

	//delete content;
}
