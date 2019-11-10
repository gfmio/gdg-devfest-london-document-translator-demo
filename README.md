# Document Translator with Google Cloud Translation

This repository is the source code accompanying my talk "Building a document
translator with Google Cloud Translation" at the GDG London DevFest 2019 on Sat,
16 Nov 2019 (10:45am - 11:10am).

It contains the C# Azure Functions API in `api` and a next.js web app as a
sample client in `web`. The service is capable of translating .docx Word
documents, .xlsx Excel documents and .pptx PowerPoint presentations using any
language supported by Google Cloud Translation.

## Instructions

### API

- To run the API, you need to have .NET Core 2 (available
  [here](https://dotnet.microsoft.com/download/dotnet-core/2.2)) and the Azure
  Functions Core Tools (instructions
  [here](https://docs.microsoft.com/en-us/azure/azure-functions/functions-run-local))
  installed.
- You also need to have a Google Cloud credentials for a project with the Google
  Cloud Translation API activated. If you need to set up new credentials, follow
  the instructions in the
  [documentation](https://cloud.google.com/translate/docs/basic/setup-basic).
- Make a copy `api/local.settings.json.template` to `api/local.settings.json`.
  The variable `GOOGLE_APPLICATION_CREDENTIALS` needs to be set to the path
  where your credentials file is located. If you created new credentials in the
  previous step, you can move the file to `api/credentials.json` and it will
  work out of the box.
- To run the API, execute `func host start` in the `api` directory. The API will
  launch at <http://0.0.0.0:7071>.

### Web client

- To run the web app, you need to have node.js and yarn installed.
- In the `web` directory, execute `yarn install` to install all dependencies.
- Then, execute `yarn dev` to build and start the web app, which will launch at
  <http://0.0.0.0:3000>.

## Related work

The code for processing the Word, Document is based on Microsoft's
[DocumentTranslator](https://github.com/MicrosoftTranslator/DocumentTranslator)
open source sample project, but adapted to run in a Function-as-a-Service
environment and using Google Cloud Translation to perform the translation
itself.
