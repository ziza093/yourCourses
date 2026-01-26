# yourCourses
A small app which takes the uni table and extract only the courses your class should attend.

To run the program you need to clone the repository and have Docker installed on your computer.

HOW TO RUN:

1. Build the docker image
```Docker
docker build -t `yourname`
```
2. Run the docker image
```Docker
docker run --rm -it -v ${PWD}:/app `yourname`
```

Now some new files should apper in the local repository directory. The parsed table is named *table.xlsx*.

! Feel free to leave any feedback on any problems about the installing phase or running phase at: diaconu.gabriel@proton.me