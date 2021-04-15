# Bridge
This is a windows APP developed for manage the bridge engineering project.
Related Word: winform C# database mysql Excel OLEDB Microsoft.Office.Interop.Excel

## Table of Contents

- [Background](#background)
- [Input](#Input)
- [MovieAPI](#MovieAPI)
- [Demo](#Demo)
- [Contributing](#contributing)
- [License](#license)

## Background
For now, the bridge engineering project need to deal with lots of raw data. We choose to develop a windows software in order to show and process the raw data the bridge engineering project.
In this project we finished the frame of the software which include input excel file, output excel file, search, show the raw data and some details.

## Input
You could find the code at [NLUbasedRasa.py](https://github.com/vegetablesB/MovieBot/blob/master/NLUbasedRasa.py).  
In this part, I realized extracting user intent and entity. It's easy to extract entity which is common such as year like "2020". But It's hard to extract movie name such as The Godfather or Cathch Me If You Can. After thinking, I used regular expressions to improve the ability. You can find the code at [data/cnrasa.json](https://github.com/vegetablesB/MovieBot/blob/master/data/cnrasa.json).  
I also found sometimes you cannot extract basic intent or entity because the data used to train is not enough. Maybe other data way too much. So I balanced the training data.

## MovieAPI
Thanks for [rapidapi](https://rapidapi.com/).  
And I used [IMDbapi](https://rapidapi.com/apidojo/api/imdb8?endpoint=apiendpoint_dad99933-4241-43f0-b4f2-529d652dcc96) to realize movie information search including movie name, genre, images and posters.  
You could find the code in [MovieAPI.py](https://github.com/vegetablesB/MovieBot/blob/master/MovieAPI.py).

## Demo
You could find the vedio in [demo](https://github.com/vegetablesB/MovieBot/blob/master/demo).

## Related Efforts
Thanks for Rasa spacy rapidapi and [python-telegram-bot](https://github.com/python-telegram-bot/python-telegram-bot).

## Contributing
[@vegetablesB](https://github.com/vegetablesB)

## License
[MIT © Richard McRichface.](../LICENSE)

![image](https://user-images.githubusercontent.com/44360183/114896539-03484f00-9e43-11eb-8a1c-d1c097f5b498.png)
![image](https://user-images.githubusercontent.com/44360183/114896566-0a6f5d00-9e43-11eb-9333-7e845d2b452b.png)
![image](https://user-images.githubusercontent.com/44360183/114896574-0cd1b700-9e43-11eb-8ee6-5aa49ebafc2a.png)


