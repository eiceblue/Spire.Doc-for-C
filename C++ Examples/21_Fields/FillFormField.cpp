#include <algorithm>
#include "pch.h"
#include "tinyxml2.h"

using namespace Spire::Doc;
using namespace tinyxml2;

wstring  string2wstring(string str)
{
	std::string strLocale = setlocale(LC_ALL, "");
	const char* chSrc = str.c_str();
	size_t nDestSize = mbstowcs(NULL, chSrc, 0) + 1;
	wchar_t* wchDest = new wchar_t[nDestSize];
	wmemset(wchDest, 0, nDestSize);
	mbstowcs(wchDest, chSrc, nDestSize);
	std::wstring wstrResult = wchDest;
	delete[] wchDest;
	setlocale(LC_ALL, strLocale.c_str());
	return wstrResult;
}

string wstring2string(const std::wstring& wstr)
{
	std::string result;
	result.reserve(wstr.size());
	for (size_t i = 0; i < wstr.size(); ++i)
	{
		result += static_cast<char>(wstr[i] & 0xFF);
	}
	return result;
}

wstring Trim(const std::wstring& str)
{
	auto first = find_if_not(str.begin(), str.end(), [](wint_t c) {return iswspace(c); });
	auto last = find_if_not(str.rbegin(), str.rend(), [](wint_t c) {return iswspace(c); }).base();
	return (first >= last) ? L"" : wstring(first, last);
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/FillFormField.doc";
	std::wstring inputFile = DATAPATH"/FillFormField.doc";

	//Open a Word document with form.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Load data.
	tinyxml2::XMLDocument* xpathDoc = new tinyxml2::XMLDocument();
	std::wstring wpath = DATAPATH"/User.xml";
	xpathDoc->LoadFile(wstring2string(wpath).c_str());
	tinyxml2::XMLElement* user = xpathDoc->FirstChildElement("user");

	//Fill data.
	int formFieldsCount = document->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetCount();

	for (int i = 0; i < formFieldsCount; i++)
	{
		intrusive_ptr<FormField> field = document->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetItem(i);

		tinyxml2::XMLElement* propertyNode = user->FirstChildElement(wstring2string(Trim(field->GetName())).c_str());
		if (propertyNode != nullptr)
		{
			switch (field->GetType())
			{

			case FieldType::FieldFormTextInput:
				field->SetText(string2wstring(propertyNode->GetText()).c_str());
				break;

			case FieldType::FieldFormDropDown:
			{
				intrusive_ptr<DropDownFormField> combox = Object::Dynamic_cast<DropDownFormField>(field);
				for (int j = 0; j < combox->GetDropDownItems()->GetCount(); j++)
				{
					if (combox->GetDropDownItems()->GetItem(j)->GetText() == string2wstring(propertyNode->GetText()))
					{
						combox->SetDropDownSelectedIndex(j);
						break;
					}
					if (wcscmp(field->GetName(), L"country") == 0 && wcscmp(combox->GetDropDownItems()->GetItem(j)->GetText(), L"Others") == 0)
					{
						combox->SetDropDownSelectedIndex(j);
					}
				}
				break;

			}
			default:
				//Nothing to do
				break;
			case FieldType::FieldFormCheckBox:
				std::string boolStr = propertyNode->GetText();
				bool value;
				std::istringstream(boolStr) >> boolalpha >> value;
				if (value)
				{
					intrusive_ptr<CheckBoxFormField> checkBox = Object::Dynamic_cast<CheckBoxFormField>(field);
					checkBox->SetChecked(true);
				}
				break;

			}
		}
	}

	delete xpathDoc;


	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Doc);
	document->Close();
}