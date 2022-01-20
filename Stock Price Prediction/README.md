## About
This project involves the use of Recurrent Neural Network(RNN) in the form of Long Short-Term Memory(LSTM) to predict the stock prices of Google.
## About the dataset
We are using Google’s Stock price from 5 years till now from a financial website (Yahoo Finance). The idea of this project was based on a project by students from Stanford (Financial Market Time Series Prediction with Recurrent Neural Networks — Bernal, Fok, Pidaparthi). The team used an Echo State Network instead of an LSTM. We will use their findings as a comparison to how our LSTM performs. The team trained their model from late 2004 to early 2009 with data from Yahoo Finance.
We will also train our LSTM on 5 years of data. We can see that their predictions are quite close to the actual Stock Price. We can try to get the same accuracy from our model as well.
We assume that the present day is January 01, 2017. We will then get the Google Stock price for the previous 5 years. Once we train our LSTM, we will try to predict the stock price for the month of January 2017.
