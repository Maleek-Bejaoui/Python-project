# Python-project

This code was developed by Asma Rhaiem, Montasar Borgie, and myself, Malek Bejaoui, as part of a Python mini-project at the National School of Electronics and Telecommunications of Sfax. We worked together to create this script that uses the Mediapipe library to detect hand movements in real-time video captured from a webcam and uses these gestures to control a PowerPoint presentation. We also added comments in the code to facilitate understanding and modification of the code for future users.

This code is a Python script that uses the Mediapipe library to detect and track hand movements in real-time video captured from a webcam. The script detects when the user makes a fist or an open hand gesture and uses these gestures to navigate through a PowerPoint presentation.

The script begins by importing the necessary libraries including Mediapipe, OpenCV, NumPy, and PowerPoint. It then initializes the Mediapipe library for hand detection and drawing the landmarks on the hands. The script also initializes variables to track the user's hand gestures, the current slide of the PowerPoint presentation, and the number of slides in the presentation.

The main loop of the script reads in each frame from the webcam and processes it using Mediapipe to detect hand landmarks. If hands are detected, the landmarks are drawn on the image using the Mediapipe library. The script then calculates the distance between the thumb and index finger landmarks to determine whether the user's hand is open or closed. If the hand is closed, the script advances to the next slide in the presentation, and if the hand is open, it goes back to the previous slide.

The script displays the video feed with the hand landmarks drawn on it and waits for the user to press the 'q' key to exit. Finally, it releases the video capture and closes the PowerPoint presentation, which has been updated based on the user's hand gestures.

Overall, this script demonstrates how to use computer vision techniques to detect hand gestures and use them for controlling other applications, such as PowerPoint.

