#include "pch.h"
#include <fstream>
#include <locale>
#include <codecvt>

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/GetFieldText.txt";
	std::wstring inputFile = DATAPATH"/SampleB_1.docx";

	std::wstring stringBuilder;

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get all fields in document
	intrusive_ptr<FieldCollection> fields = document->GetFields();
	for (int i = 0; i < fields->GetCount(); i++)
	{
		intrusive_ptr<Field> field = fields->GetItem(0);
		//Get field text
		std::wstring fieldText = field->GetFieldText();
		stringBuilder.append(L"The field text is \"" + fieldText + L"\".\n");
	}

	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << stringBuilder;
	write.close();

	document->Close();
}