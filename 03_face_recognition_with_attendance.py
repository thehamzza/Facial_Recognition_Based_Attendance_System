import cv2
import numpy as np
import os
import xlsxwriter
import numpy as np
#to get unique values from attendnace list
def unique(list1):
    x = np.array(list1)

from datetime import datetime
now = datetime.now()
current_date_time = now.strftime("%d-%m-%Y %H-%M-%S")


# Create a workbook and add a worksheet for each session of attendance.
workbook = xlsxwriter.Workbook(f'C:\\Users\\Softsys\\Desktop\\Python\\FacialRecognitionProject\\Attendance\\Attendance of {str(current_date_time)}.xlsx')

worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Serial No.')
worksheet.write('B1', 'Name')
worksheet.write('C1', 'Status')


recognizer = cv2.face.LBPHFaceRecognizer_create()
recognizer.read('C:/Users/Softsys/Desktop/Python/FacialRecognitionProject/trainer/trainer.yml')
cascadePath = "C:/Users/Softsys/Desktop/Python/FacialRecognitionProject/Cascades/haarcascade_frontalface_default.xml"
faceCascade = cv2.CascadeClassifier(cascadePath);

font = cv2.FONT_HERSHEY_SIMPLEX

# iniciate id counter
id = 0

# names related to ids: example ==> Marcelo: id=1,  etc
names = ['M Hamza S.','Haris J.', 'Abdullah S.', 'Saad S.']
present_students=[]
absent_students=[]

# Initialize and start realtime video capture
cam = cv2.VideoCapture(0)
cam.set(3, 640)  # set video widht
cam.set(4, 480)  # set video height

# Define min window size to be recognized as a face
minW = 0.1 * cam.get(3)
minH = 0.1 * cam.get(4)

while True:

    ret, img = cam.read()
    img = cv2.flip(img, 1)  # Flip vertically

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    faces = faceCascade.detectMultiScale(
        gray,
        scaleFactor=1.2,
        minNeighbors=5,
        minSize=(int(minW), int(minH)),
    )

    for (x, y, w, h) in faces:

        cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)

        id, confidence = recognizer.predict(gray[y:y + h, x:x + w])

        # Check if confidence is less then 100 ==> "0" is perfect match
        if (confidence < 100):
            id = names[id]
            new_confidence = "  {0}%".format(round(100 - confidence))

            #writing attendance to excel sheet
            #only consider those students present whose confidence is more than 50%
            attendance_confidence="  {0}".format(round(100 - confidence))
            if (int(attendance_confidence)>50):
                #writing present students list
                present_students.append(id)

        elif (confidence > 100):
            id = "unknown"
            new_confidence = "  {0}%".format(round(100 - confidence))

        cv2.putText(img, str(id), (x + 5, y - 5), font, 1, (255, 255, 255), 2)
        # print("Name of Student : " + str(id))
        # print("Confidence : " + str(confidence))
        cv2.putText(img, str(new_confidence), (x + 5, y + h - 5), font, 1, (255, 255, 0), 1)

    cv2.imshow('camera', img)

    k = cv2.waitKey(10) & 0xff  # Press 'ESC' for exiting video
    if k == 27:
        break


present_students = np.unique(present_students)


#differnce of present and absent students list from original dataset
#z=x xor y ; elements of x minus y
absent_students = set(names) ^ set(present_students)
absent_students=list(absent_students)
print("\nPresent Students : ", present_students)
print("\nAbsent Students : ", absent_students)

#writing attenance to excel sheet
row = 1
col = 0
count=0
for i in range (len(present_students)):
    #serial no.
    worksheet.write(row, col,count+1)
    #names
    worksheet.write(row, col + 1, present_students[i])
    #status
    worksheet.write(row, col + 2, "P")
    row += 1
    count += 1

for i in range (len(absent_students)):
    worksheet.write(row, col,count+1)
    worksheet.write(row, col + 1, absent_students[i])
    worksheet.write(row, col + 2, "A")
    row += 1
    count +=1


workbook.close()
print("\n [INFO] Attendance Taken and Saved in Excel Sheet :) ")
# Do a bit of cleanup
print("\n [INFO] Exiting Program and cleanup stuff")
cam.release()
cv2.destroyAllWindows()