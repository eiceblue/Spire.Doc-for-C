#include "pch.h"
using namespace Spire::Doc;

void SplitDocIntoMultipleHtml(const wstring& input, const wstring& outDirectory);
bool IsInNextDocument(intrusive_ptr<DocumentObject> element);

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SplitDocIntoHtmlPages.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SplitDocIntoHtmlPages/";

	//Split a document into multiple html pages.
	SplitDocIntoMultipleHtml(inputFile, outputFile);

}

void SplitDocIntoMultipleHtml(const wstring& input, const wstring& outDirectory)
{
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(input.c_str());

	intrusive_ptr<Document> subDoc = nullptr;
	bool first = true;
	int index = 0;
	int sectionCount = document->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(i);
		int secChildObjectsCount = sec->GetBody()->GetChildObjects()->GetCount();
		for (int j = 0; j < secChildObjectsCount; j++)
		{
			intrusive_ptr<DocumentObject> element = sec->GetBody()->GetChildObjects()->GetItem(j);
			if (IsInNextDocument(element))
			{
				if (!first)
				{
					//Embed css tyle and image data into html page
					subDoc->GetHtmlExportOptions()->SetCssStyleSheetType(CssStyleSheetType::Internal);
					subDoc->GetHtmlExportOptions()->SetImageEmbedded(true);
					//Save to html file
					wstring filePath = outDirectory + L"out-" + to_wstring(index++) + L".html";
					subDoc->SaveToFile(filePath.c_str(), FileFormat::Html);
					subDoc = nullptr;
				}
				first = false;
			}
			if (subDoc == nullptr)
			{
				subDoc = new Document();
				subDoc->AddSection();
			}
			subDoc->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add(element->Clone());
		}
	}
	if (subDoc != nullptr)
	{
		//Embed css tyle and image data into html page
		subDoc->GetHtmlExportOptions()->SetCssStyleSheetType(CssStyleSheetType::Internal);
		subDoc->GetHtmlExportOptions()->SetImageEmbedded(true);
		//Save to html file
		wstring filePath = outDirectory + L"out-" + to_wstring(index++) + L".html";
		subDoc->SaveToFile(filePath.c_str(), FileFormat::Html);
		subDoc->Close();
	}
}

bool IsInNextDocument(intrusive_ptr<DocumentObject> element)
{

	if (Object::CheckType<Paragraph>(element))
	{
		intrusive_ptr<Paragraph> p = boost::dynamic_pointer_cast<Paragraph>(element);
		if (wcscmp(p->GetStyleName(), L"Heading1") == 0)
		{
			return true;
		}
	}
	return false;
}
