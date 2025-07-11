
#include "../pch.h"


using namespace std;
using namespace Spire::Doc;
using namespace Spire::Doc::Pages;


int main()
{
	// Specify the file path
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"in.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"out.txt";

	// Create a new instance of Document
	intrusive_ptr<Document> document = new Document();

	//Load the document from the specified file
	document->LoadFromFile(inputFile.c_str(), FileFormat::Docx);
	intrusive_ptr<FixedLayoutDocument> layoutDoc = new FixedLayoutDocument(document);
	wstring result;

	// Create a FixedLayoutDocument object using the loaded document
	intrusive_ptr<FixedLayoutLine> line = layoutDoc->GetPages()->GetItem(0)->GetColumns()->GetItem(0)->GetLines()->GetItem(0);
	result.append(L"Line: ");
	result.append(line->GetText());
	result.append(L"\n");

	// Retrieve the original paragraph associated with the line
	intrusive_ptr<Paragraph> para = line->GetParagraph();
	result.append(L"Paragraph text: ");
	result.append(para->GetText());
	result.append(L"\n");

	// Retrieve all the text that appears on the first page in plain text format (including headers and footers).
	wstring pageText = layoutDoc->GetPages()->GetItem(0)->GetText();
	result.append(pageText);
	result.append(L"\n");

	// Loop through each page in the document and print how many lines appear on each page.
	for (int i = 0; i < layoutDoc->GetPages()->GetCount(); i++)
	{
		intrusive_ptr<FixedLayoutPage> page = layoutDoc->GetPages()->GetItem(i);
		intrusive_ptr<LayoutCollection> lines = page->GetChildEntities(LayoutElementType::Line, true);
		result.append(L"Page ");
		result.append(std::to_wstring(page->GetPageIndex()));
		result.append(L" has ");
		result.append(std::to_wstring(lines->GetCount()));
		result.append(L" lines.");
		result.append(L"\n");
	}

	// Perform a reverse lookup of layout entities for the first paragraph
	result.append(L"\n");
	result.append(L"The lines of the first paragraph:");
	result.append(L"\n");
	intrusive_ptr<Paragraph> para2 = (Object::Dynamic_cast<Section>(document->GetFirstChild()))->GetBody()->GetParagraphs()->GetItemInParagraphCollection(0);
	intrusive_ptr<LayoutCollection> paragraphLines = layoutDoc->GetLayoutEntitiesOfNode(para2);
	for (int i = 0; i < paragraphLines->GetCount(); i++)
	{
		intrusive_ptr<FixedLayoutLine> paragraphLine = Object::Dynamic_cast<FixedLayoutLine>(paragraphLines->GetItem(i));
		result.append(paragraphLine->GetText());
		result.append(L"\n");
		result.append(paragraphLine->GetRectangle()->ToString());
		result.append(L"\n");
		result.append(L"\n");
	}

	// Write the extracted text to a file
	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << result;
	write.close();

	// Dispose of the document resources
	document->Dispose();
}





