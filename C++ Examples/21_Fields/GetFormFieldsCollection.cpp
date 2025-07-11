#include "pch.h"
#include <fstream>
#include <locale>
#include <codecvt>

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/GetFormFieldsCollection.txt";
	std::wstring inputFile = DATAPATH"/FillFormField.doc";

	std::wstring stringBuilder;

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	intrusive_ptr<FormFieldCollection> formFields = section->GetBody()->GetFormFields();

	stringBuilder.append(L"The first section has " + std::to_wstring(formFields->GetCount()) + L" form fields.");

	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << stringBuilder;
	write.close();

	document->Close();
}
