# A task from the second laboratories of Programming languages on .NET

The project contains a simple EPPlus console application:
- recursive "ls"/"dir" with basic exception handling,
- gets extension, size and attributes of the files,
- writes the data into .xslx worksheet (1 line per file/directory with coresponding data),
- groups the lines by the directories,
- gets the 10 largest files and puts them into another ranked worksheet,
- creates percentage charts on the second worksheet (by size and count of each extension)

The app takes two arguments:
- pathToDirectory e.g. "C:\Downloads"
- depthOfTheSearch e.g. "3"
