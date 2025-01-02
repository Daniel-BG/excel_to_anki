# Excel to Anki

Convert excel files, with text and images, to anki cards.

In the excel file (in .xlsx format), have the names in one column, and the images wherever else (the script will look for them as long as they are in the same row). 

The anki cards are composed of the text of the selected column, and a reference to the image. The name of the image is in the format of `image_<text>.jpg`.

## Usage:
* `-i <input>`: Required. Set the input excel file
* `-o <output>`: Name of output card file. Default is `anki_cards.tsv` change if desired.
* `-d <dir>`: Name of output directory. Default is `anki_images`.
* `-c <col>`: Required Column index where to take the card name from. 
* `-s <val>`: Number of rows to skip before generating. Default is 1 to skip the (usually present) header. Change if desired.

## Example

First, convert the excel file:
```
python3 excel_to_anki.py -i file.xlsx -c 1
```
You will get two outputs. The `anki_cards.tsv` file and the `anki_images` folder.

Import the cards into anki and assing to a deck (from the `Import File` button). Then copy the contents of the media folder to the media folder of your anki installation (`tools> check media> View Files`). Then hit `tools> check database` to rebuild the database, and the new deck, along with the images, should be good to go!