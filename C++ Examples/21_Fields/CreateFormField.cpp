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

void SetPage(intrusive_ptr<Section> section)
{
	//the unit of all measures below is point, 1point = 0.3528 mm
	section->GetPageSetup()->SetPageSize(PageSize::A4());
	section->GetPageSetup()->GetMargins()->SetTop(72.0f);
	section->GetPageSetup()->GetMargins()->SetBottom(72.0f);
	section->GetPageSetup()->GetMargins()->SetLeft(89.85f);
	section->GetPageSetup()->GetMargins()->SetRight(89.85f);
}

void InsertHeaderAndFooter(intrusive_ptr<Section> section)
{
	//insert picture and text to header
	intrusive_ptr<Paragraph> headerParagraph = section->GetHeadersFooters()->GetHeader()->AddParagraph();
	intrusive_ptr<DocPicture> headerPicture = headerParagraph->AppendPicture(DATAPATH"/Header.png");
	//header text
	intrusive_ptr<TextRange> text = headerParagraph->AppendText(L"Demo of Spire.Doc");
	text->GetCharacterFormat()->SetFontName(L"Arial");
	text->GetCharacterFormat()->SetFontSize(10);
	text->GetCharacterFormat()->SetItalic(true);
	headerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//border
	headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::Single);
	headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetSpace(0.05F);


	//header picture layout - text wrapping
	headerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);

	//header picture layout - position
	headerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
	headerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	headerPicture->SetVerticalOrigin(VerticalOrigin::Page);
	headerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Top);

	//insert picture to footer
	intrusive_ptr<Paragraph> footerParagraph = section->GetHeadersFooters()->GetFooter()->AddParagraph();

	intrusive_ptr<DocPicture> footerPicture = footerParagraph->AppendPicture(DATAPATH"/Footer.png");

	//footer picture layout
	footerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);
	footerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
	footerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	footerPicture->SetVerticalOrigin(VerticalOrigin::Page);
	footerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);

	//insert page number
	footerParagraph->AppendField(L"page number", FieldType::FieldPage);
	footerParagraph->AppendText(L" of ");
	footerParagraph->AppendField(L"number of pages", FieldType::FieldNumPages);
	footerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//border
	footerParagraph->GetFormat()->GetBorders()->GetTop()->SetBorderType(BorderStyle::Single);
	footerParagraph->GetFormat()->GetBorders()->GetTop()->SetSpace(0.05F);
}

void AddTitle(intrusive_ptr<Section> section)
{
	intrusive_ptr<Paragraph> title = section->AddParagraph();
	intrusive_ptr<TextRange> titleText = title->AppendText(L"Create Your Account");
	titleText->GetCharacterFormat()->SetFontSize(18);
	titleText->GetCharacterFormat()->SetFontName(L"Arial");
	titleText->GetCharacterFormat()->SetTextColor(Color::FromArgb(0x00, 0x71, 0xb6));
	title->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	title->GetFormat()->SetAfterSpacing(8);
}

void AddFormDemo(intrusive_ptr<Section> section)
{
	intrusive_ptr<ParagraphStyle> descriptionStyle = new ParagraphStyle(section->GetDocument());
	descriptionStyle->SetName(L"description");
	descriptionStyle->GetCharacterFormat()->SetFontSize(12);
	descriptionStyle->GetCharacterFormat()->SetFontName(L"Arial");
	descriptionStyle->GetCharacterFormat()->SetTextColor(Color::FromArgb(0x00, 0x45, 0x8e));
	section->GetDocument()->GetStyles()->Add(descriptionStyle);

	intrusive_ptr<Paragraph> p1 = section->AddParagraph();
	wstring* text1 = new wstring();
	text1->append(L"So that we can verify your identity and find your information, ");
	text1->append(L"please provide us with the following information. ");
	text1->append(L"This information will be used to create your online account. ");
	text1->append(L"Your information is not public, shared in anyway, or displayed on this site");
	p1->AppendText(text1->c_str());
	p1->ApplyStyle(descriptionStyle->GetName());

	delete text1;

	intrusive_ptr<Paragraph> p2 = section->AddParagraph();
	std::wstring text2 = L"You must provide a real email address to which we will send your password.";
	p2->AppendText(text2.c_str());
	p2->ApplyStyle(descriptionStyle->GetName());
	p2->GetFormat()->SetAfterSpacing(8);

	//field group label style
	intrusive_ptr<ParagraphStyle> formFieldGroupLabelStyle = new ParagraphStyle(section->GetDocument());
	formFieldGroupLabelStyle->SetName(L"formFieldGroupLabel");
	formFieldGroupLabelStyle->ApplyBaseStyle(L"description");
	formFieldGroupLabelStyle->GetCharacterFormat()->SetBold(true);
	formFieldGroupLabelStyle->GetCharacterFormat()->SetTextColor(Color::GetWhite());
	section->GetDocument()->GetStyles()->Add(formFieldGroupLabelStyle);

	//field label style
	intrusive_ptr<ParagraphStyle> formFieldLabelStyle = new ParagraphStyle(section->GetDocument());
	formFieldLabelStyle->SetName(L"formFieldLabel");
	formFieldLabelStyle->ApplyBaseStyle(L"description");
	formFieldLabelStyle->GetParagraphFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);
	section->GetDocument()->GetStyles()->Add(formFieldLabelStyle);

	//add table
	intrusive_ptr<Table> table = section->AddTable();

	//2 columns of per row
	table->SetDefaultColumnsNumber(2);

	//default height of row is 20point
	table->SetDefaultRowHeight(20);

	//load form config data

	/*Stream stream = File::OpenRead(DataPath"/Demo/Form.xml");
	XPathintrusive_ptr<Document> xpathDoc = new XPathDocument(stream);*/
	tinyxml2::XMLDocument* xpathDoc = new tinyxml2::XMLDocument();
	std::wstring wpath = DATAPATH"/Form.xml";
	std::string finalPath = wstring2string(wpath);
	xpathDoc->LoadFile(finalPath.c_str());
	tinyxml2::XMLElement* root = xpathDoc->RootElement();
	std::vector<tinyxml2::XMLElement*> sectionNodes;
	const char* target = "section";
	for (tinyxml2::XMLElement* tempSetion = root->FirstChildElement(target); tempSetion; tempSetion = tempSetion->NextSiblingElement(target))
	{
		sectionNodes.push_back(tempSetion);
	}
	for (tinyxml2::XMLElement* node : sectionNodes)
	{
		//create a row for field group label, does not copy format
		intrusive_ptr<TableRow> row = table->AddRow(false);
		row->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::FromArgb(0x00, 0x71, 0xb6));
		row->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);

		//label of field group
		intrusive_ptr<Paragraph> cellParagraph = row->GetCells()->GetItemInCellCollection(0)->AddParagraph();
		std::wstring wideCStr = string2wstring(node->Attribute("name"));
		cellParagraph->AppendText(wideCStr.c_str());
		cellParagraph->ApplyStyle(formFieldGroupLabelStyle->GetName());

		//XPathNodeIterator* fieldNodes = node->Select(L"field");
		std::vector<tinyxml2::XMLElement*> fieldNodes;
		const char* fieldStr = "field";
		for (tinyxml2::XMLElement* tempField = node->FirstChildElement(fieldStr); tempField; tempField = tempField->NextSiblingElement(fieldStr))
		{
			fieldNodes.push_back(tempField);
		}

		for (tinyxml2::XMLElement* fieldNode : fieldNodes)
		{
			//create a row for field, does not copy format
			intrusive_ptr<TableRow> fieldRow = table->AddRow(false);

			//field label
			fieldRow->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
			intrusive_ptr<Paragraph> labelParagraph = fieldRow->GetCells()->GetItemInCellCollection(0)->AddParagraph();
			std::wstring wideCStrLab = string2wstring(fieldNode->Attribute("label"));
			labelParagraph->AppendText(wideCStrLab.c_str());
			labelParagraph->ApplyStyle(formFieldLabelStyle->GetName());

			fieldRow->GetCells()->GetItemInCellCollection(1)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
			intrusive_ptr<Paragraph> fieldParagraph = fieldRow->GetCells()->GetItemInCellCollection(1)->AddParagraph();
			std::wstring fieldId = string2wstring(fieldNode->Attribute("id"));
			//C# TO C++ CONVERTER NOTE: The following 'switch' operated on a string and was converted to C++ 'if-else' logic:
			//					switch (fieldNode.GetAttribute("type", ""))
			//ORIGINAL LINE: case "text":
			std::wstring typeWideStr = string2wstring(fieldNode->Attribute("type"));
			if (typeWideStr == L"text")
			{
				//add text input field
				intrusive_ptr<TextFormField> field = Object::Dynamic_cast<TextFormField>(fieldParagraph->AppendField(fieldId.c_str(), FieldType::FieldFormTextInput));

				//set default text
				field->SetDefaultText(L"");
				field->SetText(L"");

			}
			//ORIGINAL LINE: case "list":
			else if (string2wstring(fieldNode->Attribute("type")) == L"list")
			{
				//add dropdown field
				intrusive_ptr<DropDownFormField> list = Object::Dynamic_cast<DropDownFormField>(fieldParagraph->AppendField(fieldId.c_str(), FieldType::FieldFormDropDown));

				//add items into dropdown.
				//XPathNodeIterator* itemNodes = fieldNode->Select("item");
				std::vector<tinyxml2::XMLElement*> itemNodes;
				const char* itemStr = "item";
				for (tinyxml2::XMLElement* tempItem = fieldNode->FirstChildElement(itemStr); tempItem; tempItem = tempItem->NextSiblingElement(itemStr))
				{
					itemNodes.push_back(tempItem);
				}

				for (tinyxml2::XMLElement* itemNode : itemNodes)
				{
					//list->GetDropDownItems()->Add(itemNode->SelectSingleNode(L"text()")->GetValue());
					list->GetDropDownItems()->Add(string2wstring(itemNode->GetText()).c_str());
				}

			}
			//ORIGINAL LINE: case "checkbox":
			else if (string2wstring(fieldNode->Attribute("type")) == L"checkbox")
			{
				//add checkbox field
				fieldParagraph->AppendField(fieldId.c_str(), FieldType::FieldFormCheckBox);
			}
		}

		//merge field group row. 2 columns to 1 column
		table->ApplyHorizontalMerge(row->GetRowIndex(), 0, 1);
	}

	delete xpathDoc;
	
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CreateFormField.doc";


	//Create Word document.
	intrusive_ptr<Document> document = new Document();
	intrusive_ptr<Section> section = document->AddSection();

	//Page setup.
	SetPage(section);

	//Insert header and footer.
	InsertHeaderAndFooter(section);

	//Add title.
	AddTitle(section);

	//Add form.
	AddFormDemo(section);

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Doc);
	document->Close();
}