# Document template filler
If you have a lot of documents of the same type, for example datasheets of a lot of devices of the same type, you can generate these documents from a pre-defined template without opening and closing every document in Word. You just have to enter the requiered values into the entry fields and generate the document.
How to make a pre-defined template: make and format the word documnet, the parts to modify/fill must be between {{ and }}, for example: {{laser_type}}. The collection of these expressions is called "context".
How to use the program:
1. Load template: load the word document.
2. Get context: extracts the parts to fill from the document and shows their keywords and the corresponding entries.
3. Choose default context: not implemented yet.
4. Render context: saves the document with the given filename. Can't insert image into the document yet.
