Teaching Computers To Read Part 3 (Numbers)
===========================================

This project uses a neural net to recognise the numbers 0 to 9.

  This project aims to create a piece of optical character recognition software.
The way I have intended to do this is by using a neural net which are good for
this kind of generalisation and pattern recognition - conventional computing
techniques fail.  The aim is to be able to teach computers to read (hence the 
name of the project) things like postcodes, text within bitmaps - and not
really worrying whether these are hand-written, use an unusual font, or use a
standard font.  The result should be the same - correct recognition of a 
character.

The code constantly refers to "glyphs".  A glyph is the shape and design of 
a particular letter, and these obviously vary between fonts.

Using The Program
=================

  Before you create your own net and train it - load the one included in the
sample - wt3.nnt
This has been trained on over 28,000 glyphs and has an good accuracy rate.

To test recognition, you can either
a) Draw into the top left of the big picture (just hold down the mouse and move), 
 and then click on "Test Glyph" to test whether the net recognises what number 
 it is.
b) Select a font and just press a key - the glyph will automatically appear and
 it's output run through the net.
c) Run the test battery (not recommended, 'cos it takes ages).  Note that if you
 end up with different test results to me on the same net (wt1.nnt), it is
 almost certainly due to the fact that you haven't got all the fonts that I
 have!

All output from the net appears in the Output frame.
For example, if you have 97.23 in the "1" box, this means that the net is 97.23%
sure that the glyph is a number one.

You can create and train your own net.  If you use a different 
Learning Coefficient, or different fonts, or different Annealing rates, and you end 
up with better results please contact me - I'm still experimenting with this, so 
any help would be greatly appreciated.
Don't worry too much if initial training is chasing the previous character - this is
a natural result of using a relatively high learning co-efficient to start with.

Any bug-fixes, suggestions, and queries to bigcalm@hotmail.com


New For Part 3
==============

Part 3 may seem very similar to part 2 on the surface, but the internal workings have been
radically altered.  Instead of using Class/Collection structures to hold the net, an array
based structure has been used instead.  This is _so_ much faster (though a lot more 
complicated to code).  Destruction of the net now takes 10 miliseconds rather than 10 
minutes!

Initialisation of the net now uses the Nguyen-Widrow algorithm - and I was stunned by the
enormous improvement this made.

Part 3 also introduces simulated annealing.  Thanks to Chikh for his help with this.
Simulated annealing is used in neural nets to gradually adjust the neural net's learning
co-efficient, and reduce eventual noise in the net.

Load/Save files have been altered slightly to accomodate simulated annealing, but any files
saved in 1.2 format will still load ok.

Gone are the array sorts, the output has been scaled (as requested by Ulli), and a few other
minor bug fixes.  CRandom has also gone (I need double random values - single isn't good enough).


Neural Nets
===========

  My neural net code was originally written by Ulli, and is now _heavily_
modified.  Neural nets are modelled on the way the brain works, and by
modelling this, we gain the generalisation and pattern recognition capabilities
which humans do so well (and, conventional computers do so badly).  This comes
at a price - neural net modelling is slow without specialised hardware, and
neural nets are prone to getting things wrong (just like us).
All neural nets require training - and in general the more training a neural
net is given the better it learns.
Neural nets are not just used for character recognition - any project that
requires generalisation and pattern recognition can be helped by using a 
neural net - applications range from speech recognition, data mining - anything
where pattern recognition can help.

The second project (prjNeuralNet.vbp) is written to be a dll that anyone can use.

Further information on neural nets can be found at:
www.faqs.org  (search for Neural).

Obligatory Disclaimer
=====================
  This source is provided on the basis that the author holds no liability for
the consequences of it's use.  It is by no means bullet-proof and shouldn't be
used in any critical application such as air-traffic control, hospitals, etc., 
etc.  If you use any major part of the code (I don't mind you stealing snippets)
please credit the author(s).

Credits
=======

I have used a lot of code from other people - my thanks, and here are the 
credits:

Ulli: Original Neural Net code.
Tom Sawyer: Font API enumeration
Roger Johansson: Anti-aliased text.


Part 4 (The next project)
=========================

 Part 4 will attempt to train on all lower case characters.
All of Ulli's original classes (the slow Class/Collection based ones) will be 
removed and unsupported from here on (though credits will remain).

  - Jonathan Daniel, 02/05/01
