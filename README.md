# Optimized-Shopping-Schedule
Instructions Before Using the Application:
Create and save an excel file on your computer with a list of last names.
-> It would be beneficial if the list of last names are store customers or people who live within the area. (Our recommendation would be to use the last names of any people signed up with your stores rewards program)
-> In the excel file, create a column with a header as “Last Name”, then put the last names in each cell in that column.
Download latest version of python at https://www.python.org/downloads/
Double click on the python (.py) file
An instructions file will be saved onto your computer. Open the file called “Instructions.txt” for step-by-step instructions on how to use the code.


Why The Application Is Helpful:

Since the COVID-19 pandemic outbreak, people have been isolating themselves at home to try and slow the spread of the virus. The one thing that everyone still needs to do is go shopping to buy essentials. People tend to go shopping around the same few hours every day. This results in there being hours of over-crowding and hours of under-crowding. Because of this, the virus is still capable of spreading within the stores.

Some stores have taken necessary precautions by only allowing a set number of people within the stores at once. However, people who arrive at the store usually have to wait a long time in a line before they can even go inside.

We have come up with a solution to help slow the spread of the virus that allows people to not have to wait to go shopping. Our code creates an optimized shopping schedule based on people’s last names. This way, everyone will have a designated shopping time slot throughout the week. With this solution, people will have to plan in advance when they will be going shopping, but they can do so by checking which time they are supposed to go. As an example, from 9am to 11am, anyone with a last name that begins with A, B, C, D, or E, may go shopping.

You will be asked to input your store’s opening hours, as well as how many groups you would like the last names data split up into. If you would like more groups with less people in each group, or less groups with more people in each group, that is up to you.

It is your choice to have the schedule printed on a PDF file (.pdf), a Microsoft Excel file (.xlsx), or both. This will be asked of you while using the code.


What Cases The Code Can Handle:

Any store hours can be used (unless open from a pm time to an am time).
The store hours can be different on a Saturday or Sunday than on a weekday.
Any number of groups less than or equal to 10 can be used.


How the Algorithm Works:
 
Reads an excel file with last names in it.
Creates groups that represent alphabetical letters. (Ex: Group 1: A-D, Group 2: E-K, etc.)
Finds the best group distribution possible by making the number of people within each group to be the closest possible. (Ex: Group1 = 10 people, Group2 = 11 people, Group3 = 9 people, is better than Group1 = 15 people, Group2 = 3 people, Group3 = 12 people)
The best group allocation is found by calculating the minimum statistical variance of the possible groups.


About Us:

The people who created this algorithm are Bowen Zhang and Jordan Pivato. They are students at the University of Waterloo studying computer science and mathematics.

They spent countless hours every day for the past month creating this project. Everyone needs to play a part in trying to help stop this virus. Some people can do that by working in a hospital, some by donating to charities, some by helping give food to people in need. They decided that they could help by creating a schedule that will reduce the amount of human interaction that is needed within stores.

They noticed a problem that needs to be solved and wanted to do their best to try and help. They ask for nothing in return other than consideration to implement this shopping schedule into the stores.

Since this is free to use, you can try it out and see if you like the results before implementing anything.

You can make it mandatory for people by possibly making them show ID before entering the store (to prove their last name). You can also just make the schedule a suggestion for people to follow. How you implement the schedule is totally up to you based on what you think will work the best.
