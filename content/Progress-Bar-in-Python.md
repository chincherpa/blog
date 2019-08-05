Title: Progress Bar in Python
Date: 2019-06-26 12:21
Author: Lulef
Category: Sammlung
Slug: progressbar-in-python
Status: published

<https://codingdose.info/2019/06/15/how-to-use-a-progress-bar-in-python/>

## How to Easily Use a Progress Bar in Python 


![alt](https://codingdose.info/assets/images/progressbars/carbon.png)

We’ve all been there, your code is performing a job and it’s going to take a while. I’m an impatient guy, so it would be nice to have an ETA or a progress bar to show us. Fortunately, there are libraries out there than can help us to achieve this!

There’s two ways in which we can integrate a progress bar into our loops, via a [*Context Manager*](https://jeffknupp.com/blog/2016/03/07/python-with-context-managers/) or just wrapping up an iterable object into a method.

We’re going to be testing Progress, ProgressBar2, TQDM, Click and Clint, so make sure to create your testing environment with [Pipenv](https://codingdose.info/2018/02/20/pipenv-development-workflow/):

````
$ mkdir progressbar-testing
$ cd progressbar-testing
$ pipenv install tqdm progressbar2 click clint
````

[alt](https://codingdose.info/2019/06/15/how-to-use-a-progress-bar-in-python/#Progress "Progress")
[Progress](https://github.com/verigak/progress/) {#Progress}

While I was testing each library, this is one I really liked, and this is because it has a lot of progress bar styles that you can play with. We’re not going to look at *all* of them, but you can take a look at the source code and documentation if you have further questions.

### [alt](https://codingdose.info/2019/06/15/how-to-use-a-progress-bar-in-python/#Progress-BAR "Progress - BAR")Progress - BAR {#Progress-BAR}

This is the most basic one, is basically a progress bar that is being filled by a hash. It works pretty easily with a context manager, you can use this snippet as an example:

````
from time import sleep
from progress.bar import Bar
with Bar('Processing...') as bar:
    for i in range(100):
        sleep(0.02)
        bar.next()
````

As you can see, in this case we only import the `Bar` class and we add a *label* to our progress bar and it automatically handles our loop. At the end of each iteration, we have to append the `.next()` method to our `bar` object so we can update the progress bar.

<video src="https://codingdose.info/assets/images/progressbars/progress-bar.webm" preload autoplay loop controls></video>

### [alt](https://codingdose.info/2019/06/15/how-to-use-a-progress-bar-in-python/#Progress-PixelBar "Progress - PixelBar")Progress - PixelBar {#Progress-PixelBar}

We can achieve the same with other progress bar styles, let’s try the Pixel Bar:

````
from time import sleep
from progress.bar import PixelBar
with PixelBar('Processing...') as bar:
    for i in range(100):
        sleep(0.02)
        bar.next()
````

<video src="https://codingdose.info/assets/images/progressbars/progress-pixelbar.webm" preload autoplay loop controls></video>

### [alt](https://codingdose.info/2019/06/15/how-to-use-a-progress-bar-in-python/#Progress-PixelSpinner "Progress - PixelSpinner")Progress - PixelSpinner {#Progress-PixelSpinner}

Sometimes you don’t know how long it might take to perform an operation. If this is the case, then you can use the Pixel Spinner to display a pixel spinner (duh!) without an actual *progress bar*.

````
from time import sleep
from progress.spinner import PixelSpinner
with PixelSpinner('Processing...') as bar:
    for i in range(100):
        sleep(0.06)
        bar.next()
````

<video src="https://codingdose.info/assets/images/progressbars/progress-pixelspinner.webm" preload autoplay loop controls></video>

Pretty neat right? You can read more about Progress here:

-   [Source Code](https://github.com/verigak/progress/)


[alt](https://codingdose.info/2019/06/15/how-to-use-a-progress-bar-in-python/#TQDM "TQDM")[TQDM](https://github.com/tqdm/tqdm) {#TQDM}
----------------------------------------------------------------------------------------------------------------------------------------

[TQDM](https://github.com/tqdm/tqdm), a short for *taqadum* which means *progress* in Arabic, is a very popular choice, even more for data scientist or analysts, as it provides a really fast framework with a lot of customizations and information for you to work with.

It’s also smart enough to use it with barely two lines of code. Just provide an iterable to the function `tqdm()` and your good to go.

````
from tqdm import tqdm
from time import sleep
for i in tqdm(range(100)):
    sleep(0.02)
````

And there you go! You have a lot of information on your progress bar such as a percentage, the length of your iterable, an ETA and even *iterables per seconds*!

<video src="https://codingdose.info/assets/images/progressbars/tqdm.webm" preload autoplay loop controls></video>

You can also add a label to your progress bar, displaying each object along the way:

````
import string
from tqdm import tqdm
from time import sleep
# A list from A to Z wrapped around TQDM function
progress_bar = tqdm(list(string.ascii_lowercase))
for letter in progress_bar:
    progress_bar.set_description(f'Processing {letter}...')
    sleep(0.09)
````

<video src="https://codingdose.info/assets/images/progressbars/tqdm-label.webm" preload autoplay loop controls></video>

TQDM, from my perspective, is a really important project if you’re into Data Science or need a incredibly fast way to show progress to your operations, you can read more about it here:

-   [Source Code & Documentation](https://github.com/tqdm/tqdm)
-   [Wiki](https://github.com/tqdm/tqdm/wiki)

------------------------------------------------------------------------

[alt](https://codingdose.info/2019/06/15/how-to-use-a-progress-bar-in-python/#Click "Click")[Click](https://click.palletsprojects.com/) {#Click}
-------------------------------------------------------------------------------------------------------------------------------------------------

Click is one of the best libraries out there to create a [Command Line Interface](https://en.wikipedia.org/wiki/Command-line_interface) to your apps or libraries. I cannot recommend it enough!

It also includes a very simple progress bar as an utility, you can use it inside a context manager, just like this:

````
import click
from time import sleep
# Fill character is # by default, you can change it
# for any other char you want, or even change the color.
fill_char = click.style('=', fg='yellow')
with click.progressbar(range(100), label='Loading...', fill_char=fill_char) as bar:
    for i in bar:
        sleep(0.02)
````

You can also see that we can change the color of the progress bar meter, and we also have an ETA.

<video src="https://codingdose.info/assets/images/progressbars/click.webm" preload autoplay loop controls></video>

If you want to know more about the progress bar utility in Click, you can check it out here:

-   [Click - Showing Progress Bars](https://click.palletsprojects.com/en/7.x/utils/#showing-progress-bars)

------------------------------------------------------------------------

[alt](https://codingdose.info/2019/06/15/how-to-use-a-progress-bar-in-python/#ProgressBar2 "ProgressBar2")[ProgressBar2](https://github.com/WoLpH/python-progressbar) {#ProgressBar2}
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

This is also a very popular choice and one that is easy to use. It also work with widgets to calculate the current progress, such as AbsoluteETA, AdaptiveETA, AdaptiveTransferSpeed and others which are very interesting.

It’s implementation is also very simple. As with TQDM, only need 2 lines of code are enough to get us started:

````
from time import sleep
from progressbar import progressbar
for i in progressbar(range(100)):
    sleep(0.02)
````

And yes, although you install it as *progressbar2* make sure you import it as *progressbar*. Here’s how it displays the progress bar by default.

<video src="https://codingdose.info/assets/images/progressbars/progressbar2.webm" preload autoplay loop controls></video>

You can checkout the ProgressBar2 documentation and homepage here:

-   [Homepage](https://github.com/WoLpH/python-progressbar)
-   [Documentation](https://progressbar-2.readthedocs.io/en/latest/index.html)
-   [Widgets](https://progressbar-2.readthedocs.io/en/latest/_modules/progressbar/widgets.html)


[alt](https://codingdose.info/2019/06/15/how-to-use-a-progress-bar-in-python/#Clint "Clint")[Clint](https://github.com/kennethreitz/clint) {#Clint}

And lastly we have Clint, which stands for ***C** ommand **L** ine **IN** terface **T** ools*, which is not maintained anymore, but I will show it here just to pay my respects.

Creating a progress bar is just as easy as we have seen with the other tools, this one doesn’t require a context manager. Here we can see a regular progress bar and a *Mill* style progress bar:

````
from time import sleep
from clint.textui import progress
print('Clint - Regular Progress Bar')
for i in progress.bar(range(100)):
    sleep(0.02)
print('Clint - Mill Progress Bar')
for i in progress.mill(range(100)):
    sleep(0.02)
````

And here you can see it in action:

<video src="https://codingdose.info/assets/images/progressbars/clint.webm" preload autoplay loop controls></video>

------------------------------------------------------------------------

There are other libraries out there but these are the ones that I definitely recommend you to check out. Also, here’s a snippet of code that tests all of the progress bar styles and libraries that we have tested so far. Just make sure to install the appropriate libraries.

You can clone it with Git:\

````
$ git clone https://gist.github.com/33e56c93c3c43cf70f19ecbfc921e358.git progressbar-testing
````
