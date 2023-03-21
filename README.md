# Convert And Rename Information Access Documents (CARIAD)

A tool developed to streamline a specific administrative task, according to the particular requirements of my workplace.

*Outcome:* 
* to produce an 'Information Access' pack

*Given:*
* a directory containing the documents identified for inclusion in the pack
* a spreadsheet listing the documents' current filename and a running order number ('Index Number')

*Task:*
* rename each document to the corresponding Index Number
* convert each document to pdf format (excluding Excel formats)

## Usage

Clone the repository.

Run cariad from your terminal, by running the command below and providing the path to the target directory and the name of the spreadsheet.

```bash
python cariad.py C://USER/example/filepath spreadsheet.xlsx
```

Alternatively use the GUI for a more guided experience, by running the command below and then following the prompts.

```bash
python gui.py
```

## Warnings

This process assumes the directory consists of copies of the original documents, and documents in the folder are deleted once they have been saved as the renamed pdf. To avoid deleting originals, recommend creating a copy in a separate location and running cariad there.

The process involves opening and closing some of the Office Applications (Word, Powerpoint). Recommend saving any work, and closing those applications prior to running cariad.


## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License

[MIT](https://choosealicense.com/licenses/mit/)