# TkMidpoint

**TkMidpoint** is a program that calculates the midpoints of astrological charts. There are several functions that are related to midpoints and all of them will be introduced here step by step. 

## Availability
 
The program runs on Windows, Linux and Mac operating systems.

## Dependencies

In order to run **TkAstroDb**, at least [Python](https://www.python.org/)'s 3.6 version must be installed on your computer. Note that in order to use [Python](https://www.python.org/) on the command prompt, [Python](https://www.python.org/) should be added to the PATH. There is no need to install manually the libraries that are used by the program. When the program first runs, the necessary libraries will be downloaded and installed automatically.

## Usage

**1.** Run the program by writing the below command to the console window.

**For Unix**

    python3 TkAstroDb.py

**For Windows**

    python TkAstroDb.py

**2.**  Below is the main window that is coming after the program opens.

![img1](https://user-images.githubusercontent.com/29302909/75981145-ef6ab380-5ef4-11ea-829e-77d2f89b3a66.png)

**3.** As can be seen above, there are 4 menu cascades at the top of the program. **Record** is the first menu cascade which has two options that are **Create** and **Open** menu buttons.
 
**4.** By selecting **Create** menu button, the below window occurs.

![img2](https://user-images.githubusercontent.com/29302909/75981309-4e302d00-5ef5-11ea-9ebc-9ed4c231f1a4.png)

**5.** As can be seen below, all the empty fields must be filled with the correct data in order to add a record to the database of the program. If any field is wrongly filled, the program will pop-up a Warning message. 

![img3](https://user-images.githubusercontent.com/29302909/75981504-ab2be300-5ef5-11ea-91ec-ce2ae392d91d.png)

**6.** After a record is added to the database, by selecting **Open** menu button, the records that is added to the program's database can be seen. The name of the database is **records.csv** and can be found in **TkMidpoint** directory.

![img4](https://user-images.githubusercontent.com/29302909/75981807-5341ac00-5ef6-11ea-9390-0c838c24ad58.png)

**7.** The midpoint calculations are done by using this **Open** panel. If users select a record and if users right click to the selected record, a right click menu is occurred.

![img5](https://user-images.githubusercontent.com/29302909/75982153-f7c3ee00-5ef6-11ea-98de-6b58ae58bbcc.png)

**8.** Users can edit or delete the existing records. If users click the **Export** cascade, the cascade is extended. 

![img6](https://user-images.githubusercontent.com/29302909/75982528-ec24f700-5ef7-11ea-9835-96b1a0cb1959.png)

**9.** The first cascade which name is **Midpoints** has two options.

![img7](https://user-images.githubusercontent.com/29302909/75982732-4a51da00-5ef8-11ea-81fd-1891fe12531a.png)

**10.** If users click **Natal Midpoints** button, the program calculates the natal midpoints of the selected record in less than a second. The results will be inserted in a spreadsheet file.

[Natal Midpoints](https://www.dropbox.com/s/gktow61qtrvkbs6/Tanberk_Natal%20Midpoints.ods)

**11.** If users select **Progressed Midpoint** option, the program asks for a date of progression.

![img8](https://user-images.githubusercontent.com/29302909/75983486-b2ed8680-5ef9-11ea-896a-79302c7f4bfd.png)

After writing a correct date, the midpoints of the progression are calculated in less than a second. The results will be inserted in a spreadsheet file. 

[Progressed Midpoints](https://www.dropbox.com/s/ok7zcjzkyf644q0/Tanberk_Progressed%20Midpoints.ods)

**12.** The other menu cascade is about calculating the aspects between midpoints and planets. As can be seen below, there are 5 options that users can select.

![img9](https://user-images.githubusercontent.com/29302909/75983863-5fc80380-5efa-11ea-95db-92f1a8bc6952.png)

**13.** Each option among this five produces two files. One file is a spreadsheet file that includes the record information, orb factors for each aspect, positions of midpoints that make aspects to planets, the planets that make aspect to the midpoints and the positions of that planets and also the orbs of the aspects. The second file is a text file that includes the interpretations that are valid for all these 5 options.
 
When users select options which are about **Transits** and **Progressions**, the program will ask the user to enter a date.

![img10](https://user-images.githubusercontent.com/29302909/75984097-d533d400-5efa-11ea-8a1a-7efceb131885.png)

![img8](https://user-images.githubusercontent.com/29302909/75983486-b2ed8680-5ef9-11ea-896a-79302c7f4bfd.png)

Regardless of the aspect type, the same interpretations are used for every midpoint and planet pair with an information about what the aspect type is.

[Aspects between midpoint of natal planets and natal planets](https://www.dropbox.com/s/u4rahz1qzz15u2u/Tanberk_natal_natal.ods)
[Interpretations of aspects between midpoint of natal planets and natal planets](https://www.dropbox.com/s/dz6r2m4ph1uj6j1/Tanberk_natal_natal.txt)

[Aspects between midpoint of natal planets and transit planets](https://www.dropbox.com/s/nz4l1mydnnidjcc/Tanberk_natal_transit.ods)
[Interpretations of aspects between midpoint of natal planets and transit planets](https://www.dropbox.com/s/2xw7q4a3y6n6gqb/Tanberk_natal_transit.txt)

[Aspects between midpoint of natal planets and progressed planets](https://www.dropbox.com/s/mzhyr87kxftxq2w/Tanberk_natal_progressed.ods)
[Interpretations of aspects between midpoint of natal planets and progressed planets](https://www.dropbox.com/s/0zqinx8s1vjc4ez/Tanberk_natal_progressed.txt)

[Aspects between midpoint of transit planets and natal planets](https://www.dropbox.com/s/yoozovqmail54g1/Tanberk_transit_natal.ods)
[Interpretations of aspects between midpoint of transit planets and natal planets](https://www.dropbox.com/s/papo03fxncalctj/Tanberk_transit_natal.txt)

[Aspects between midpoint of progressed planets and natal planets](https://www.dropbox.com/s/lm6o8c5iyd1icj2/Tanberk_progressed_natal.ods)
[Interpretations of aspects between midpoint of progressed planets and natal planets](https://www.dropbox.com/s/46t4too86uf78gz/Tanberk_progressed_natal.txt)

**14.** The last menu button is for calculating the midpoints of the selected record and checks whether these midpoints are making an aspect to another record that is stored in the database. After clicked that, the records which are available in the records database are listed in a treeview except from the record that is selected previously.

![img11](https://user-images.githubusercontent.com/29302909/75985124-c0f0d680-5efc-11ea-832b-a516d224f51c.png)

![img12](https://user-images.githubusercontent.com/29302909/75985226-f7c6ec80-5efc-11ea-9bfd-c00390d8676c.png)

![img13](https://user-images.githubusercontent.com/29302909/75985298-1dec8c80-5efd-11ea-962c-4b2944cdd647.png)

[Synastry Midpoints](https://www.dropbox.com/s/ruf90uasw2kp6gy/Tanberk_Flavia_Midpoint_Synastry.ods)

**15.** The second menu cascade was named as **Interpretations**. If users click the **View** menu button which is under the **Interpretations menu cascade**, a window like below occurs.

![img14](https://user-images.githubusercontent.com/29302909/75985617-ad923b00-5efd-11ea-950d-f868bbcf9f20.png)

**16.** By default, there's no available interpretation. There can be created totally **2754** possible interpretations for 18 objects. If users double click on an interpretation row, a panel like below occurs. 

![img15](https://user-images.githubusercontent.com/29302909/75985948-37da9f00-5efe-11ea-9291-7e8b449fc051.png)

Users can write their own interpretations or they can copy and paste an interpretation and write the author of the text. If the text field remained empty and users click the **Apply** button, the availability status of the midpoint planet pair doesn't change. However if users write something to the text field, then if they press **Apply** button, the status of availability changes suddenly. 
 
![img16](https://user-images.githubusercontent.com/29302909/75986094-7708f000-5efe-11ea-91b4-108f1d3359c3.png)

Users can also modify the text or author of the interpretations by double clicking the interpretation row. If the author field remains empty, it doesn't have an effect on the availability status of an interpretation. The more users add interpretations for the midpoint-planet pairs, the more the interpretation file will be filled. But by default, it's written ``None`` to the text section of the notepad file. The interpretations are stored in the **interpretations.xml** file which can be found in **TkMidpoint** directory. 

**17.** The name of third menu cascade is **Settings**. This cascade has one menu button which is for changing the orb factors.

![img17](https://user-images.githubusercontent.com/29302909/75986282-d7982d00-5efe-11ea-85c4-76fd44ff0ca7.png)

By default, the orb factors are defined as **1** degree. The valid unit is **degree-minute-second**. So users shouldn't delete the text field or shouldn't write a number such as **1.2**. The program will pop-up a warning message when the field is wrongly filled. So when users want to specify the orb factor as **+- 1.2** degree, the correct notation should be as **1° 2' 0"**.

**18.** The forth menu cascade includes two menu buttons which names are **About** and **Check for updates**. By clicking **About** menu button, users can see the contact information; and by clicking **Check for updates** menu button, users can update their scripts if any update is released.

## Licenses

TkMidpoint is released under the terms of the GNU GENERAL PUBLIC LICENSE. Please refer to the LICENSE file.
