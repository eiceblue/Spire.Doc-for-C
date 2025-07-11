#include "pch.h"
#include <fstream>
#include <locale>
#include <codecvt>

using namespace Spire::Doc;
#define stringify(name) # name

wstring getFormFieldType(FormFieldType type)
{
	switch (type)
	{
	case FormFieldType::CheckBox:
		return L"CheckBox";
		break;
	case FormFieldType::DropDown:
		return L"DropDown";
		break;
	case FormFieldType::TextInput:
		return L"TextInput";
		break;
	case FormFieldType::Unknown:
		return L"Unknown";
		break;
	}
	return L"";
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/GetFormFieldByName.txt";
	std::wstring inputFile = DATAPATH"/FillFormField.doc";

	std::wstring stringBuilder;

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	//Get form field by name
	intrusive_ptr<FormField> formField = section->GetBody()->GetFormFields()->GetItem(L"email");
	wstring formFieldName = formField->GetName();
	wstring formFieldNameType = getFormFieldType(formField->GetFormFieldType());
	stringBuilder.append(L"The name of the form field is " + formFieldName + L"\n");
	stringBuilder.append(L"The type of the form field is " + formFieldNameType);

	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << stringBuilder;
	write.close();

	document->Close();
}