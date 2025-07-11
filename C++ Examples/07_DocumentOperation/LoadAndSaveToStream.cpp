#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring outputFile = output_path + L"LoadAndSaveToStream.rtf";

	// Open the stream. Read only access is enough to load a document.
	intrusive_ptr<Stream> stream = new Stream(inputFile.c_str());

	// Load the entire document into memory.
	intrusive_ptr<Document> doc = new Document(stream);

	// You can close the stream now, it is no longer needed because the document is in memory.
	stream->Close();
	// Do something with the document

	// Convert the document to a different format and save to stream.
	intrusive_ptr<Stream> newStream = new Stream();
	doc->SaveToStream(newStream, FileFormat::Rtf);

	newStream->Save(outputFile.c_str());

	doc->Close();
}
