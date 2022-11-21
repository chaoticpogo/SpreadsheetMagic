/**
 * This sheet adds an option to the top-bar of google sheets to fill an NxM range of the
 * sheet using OpenAI's GPT-3 Language model.
 */

var API_KEY = "[API_KEY_GOES_HERE]";

var preface = "";

var NUM_TOKENS = 45;

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();

  var menuItems = [
    {name: 'Fill with GPT-3', functionName: 'gpt3fill'},
  ];

  spreadsheet.addMenu('GPT3', menuItems);
}

function _callAPI(prompt) {
  var data = {
    "model": "text-davinci-002",
    'prompt': prompt,
    'max_tokens': NUM_TOKENS,
    'temperature': 0.7,
    "top_p": 1,
    "n": 1,
    "stream": false,
    "logprobs": null,
    "stop": "\n" 
  };

  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    'payload' : JSON.stringify(data),
    'headers': {
      Authorization: 'Bearer ' + API_KEY,
    },
  };

  response = UrlFetchApp.fetch(
    'https://api.openai.com/v1/completions',
    options,
  );

  return JSON.parse(response.getContentText())['choices'][0]['text']
}

function _parse_response(response) {
  var parsed_fill = response
  /**********
  Function doesnt seem needed for content writer?
  var parsed_fill = response.slice(3);
 * parses to remove 'A: ' from the returned answer style set in the preface;
 * 
  if (parsed_fill.charAt(parsed_fill.length - 1) == '.') {
    parsed_fill = parsed_fill.slice(0, -1);
  }
 ***********/
  return parsed_fill;
}

function get_x_of_y(x, y) {
  var prompt = preface + "Q: What is the " + x + " of " + y + "?"

  var response = _callAPI(prompt);

  var parsed_response = _parse_response(response);

  return parsed_response;
}


function gpt3fill() {
  /*
  Highlight a nxm range where the leftmost column contains names of public companies.
  The header row of a column identifies properties of those companies.
  Fill in the values.
  */
  var spreadsheet = SpreadsheetApp.getActive();
  var range = spreadsheet.getActiveRange();
  var num_rows = range.getNumRows();
  var num_cols = range.getNumColumns();

  for (var x=2; x<num_cols + 1; x++) {
    article_style = range.getCell(1, x).getValue();

    for (var i=2; i<num_rows + 1; i++) {
      topic = range.getCell(i,1).getValue();
      fill_cell = range.getCell(i, x);

      result = get_x_of_y(article_style, topic);

      fill_cell.setValue([result]);
    }
  }
}

