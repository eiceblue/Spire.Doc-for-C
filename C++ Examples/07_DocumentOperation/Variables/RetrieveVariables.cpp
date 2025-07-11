#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_6.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RetrieveVariables.txt";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Retrieve name of the variable by index.
	std::wstring s1 = document->GetVariables()->GetNameByIndex(0);

	//Retrieve value of the variable by index.
	std::wstring s2 = document->GetVariables()->GetValueByIndex(0);

	//Retrieve the value of the variable by name.
	std::wstring s3 = document->GetVariables()->GetItem(L"A1");

	wstring content;
	content.append(L"The name of the variable retrieved by index 0 is: " + s1);
	content.append(L"\n");
	content.append(L"The vaule of the variable retrieved by index 0 is: " + s2);
	content.append(L"\n");
	content.append(L"The vaule of the variable retrieved by name \"A1\" is: " + s3);

	//Save to file.
	wofstream out;
	out.open(outputFile);
	out.flush();
	out << content.c_str();
	out.close();
	document->Close();
}
