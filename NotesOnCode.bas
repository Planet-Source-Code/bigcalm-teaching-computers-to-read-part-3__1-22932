Attribute VB_Name = "NotesOnCode"
Option Explicit

' This is a codeless module - just contains notes about changes.

' History of testing, changes, and musings.
' Test 1: 400*180*50*2: 14 Training Fonts.  Ok results - not recognising "O" very well in some trained fonts.
' Test 2: 400*180*50*12*2: 20 Training Fonts. Slightly better.
' Test 3: 400*180*50*12*2: 26 Training Fonts. Upside Down Character.  Great!
'    After training with 1139 glyphs I have very good accuracy rate.
'    Extremely good accuracy on fonts that it has been trained on.
'    Good accuracy on "normal looking" fonts that
'    the net has never been trained on.
'    Even some handwriting fonts have reasonable results.
'    It seems to perform poorly on fonts which are abnormally
'    wide at the moment (e.g. Wide Latin) - but including such
'    a font in the training set will aid recognition.
'    I expect with more fonts and more training cycles this will
'    improve even more.  Save: upside1.nnt
' Test 4: Use trained net from 3, and train with 29 fonts for
'    a further 500 glyphs.
'    If anything, slightly worse than before.  Save: upside2.nnt
'    It's started recognising 12-Pitch Kelly Ann Gothic with
'    high accuracy though.  HUH?  Even I can't do that...
'    Pity it's still getting some untrained basic fonts wrong tho.
'    Y'know, NNets are analogue beasts.  Anti-alias the glyphs?
'    Ick.  Later (but does need to be done).
'    Let's left-shift the glyph and run the net again.
' Test 5: 400*180*50*12*2: 29 Training fonts. Upside down & Left Shifted char.
'    Cycles: ~2000
'    Not bad at all.  Slight improvement on 3, but used more training cycles.
'    Save: upside4.nnt
' Test 6: 400*8*12*2: 1 Training font.  Upside/Left Shifted and Anti-aliased.
'    Cycles: 140
'    Does the job.  But won't generalise too well. Sigh.
'
' Bug fix.  AntiAlias text was not deleting DC's properly.  Fixed.

' Hum, thought.  Am I training on white space???  Changed so that input
' to net is always between 0 and 1, 1 being Black pixel, 0 being White Pixel. Tested Ok.

' Instead of: Draw small->Flip n Shift->Train I want to start using a constant
' height of pixels.
' So, hoping StretchBlt will help here...
' Aiming for: Draw Big -> Stretch Down to fit vertically->Flip n Shift-> Train.
' Ok. Done a test for this.  Will fail on disproportionately wide fonts such as
' Wide Latin.  Also fails on Alfredo.  I'm going to widen the input glyph.
' Ok, widened and width/height constants now fully functional.  Wide Latin still
' fails, but I can't be arsed to widen it even more - input layer is now 600 neurons.
' Rendered glyph is now looking rather good. =0)
' Training is not recommended on very narrow fonts (such as Alfredo) which
' don't scale down well.
' The AA code is a little slow.  May be removed in future, as we're scaling
' down now.
'
' Test 7: 651*8*12*2.  1 Training Font.  Fitted anti-aliased characters.
'  Crash.  Out of memory error.  I think it's this anti-alias code. Sigh.
'  NN code has run for thousands of cycles without the AA code.  Plus DCs seemed
'  to be disappearing from Windoze.  Fixed now (font deletion problem).
' A bug has just appeared in SaveNet/LoadNet - I'm going to get rid of backward
' compatibility 'cos it's getting a little too complicated in there at the moment.
' Fixed. Cycles: 130.  Ok, learning "A" & "O" from Arial quite nicely now. Save:yip2.nnt
' However, it takes ~400ms to anti-alias and ~450ms to get true textual extents.
' Which is slower than I like.
' Ok, just to check for memory loss I'll run 500 training cycles.  WooWoo! No memory loss!
' Tried to speed up true extent calculations, only to find numerous bugs. Fixed.  Saved ~100ms
' Bug fix with Font combo.
' Ok, so what now?  Let's set it up to train on lots of fonts, increase the nodes, and
' start praying....
'
' Test 8: 651*32*48*2.  25 Training Fonts.  Fitted anti-aliased characters.
' Glyphs trained: 2012.  Save: AATrain1.NNT
' I really need a "Test Battery" to test the output of the net (and gauge any
' improvements/changes I make objectively).  Ok. Test battery built.
' Here are the first objective results:
'        Trained Fonts
'        --------------------
'        Successes for A: 25
'        Failures for A: 0
'        Average Correctness for A: 99.54%
'        Average Wrongness for A: 0.33%
'        Success for O: 25
'        Failures for O: 0
'        Average Correctness for O: 99.68%
'        Average Wrongness for O: 0.37%
'
'        UnTrained Fonts
'        --------------------
'        Successes for A: 6
'        Failures for A: 1
'        Average Correctness for A: 90.34%
'        Average Wrongness for A: 11.91%
'        Success for O: 7
'        Failures for O: 0
'        Average Correctness for O: 99.46%
'        Average Wrongness for O: 1.17%
'
' Yes, that's NO failures on any known font, and ONE failure out of 14 tests on unknown.
' Not bad at all.  The one failure was Playbill "A" which has unusually large serifs.
' So, I think I can say this is the...
' ------------- < < < END OF PART 1 > > > ------------------
'
' What now? To Do list (importance in brackets):
' Need the user to be able to "draw" a glyph and then test the net with it.
'     - Whiteboard approach.  Second form.  Use Previous Real Text extents code to feed net.
'     Doesn't need to be too fancy.  Re-use someone elses code if I can get it(!) (4). Sort of done (not very well tho).
' Need to tidy up the layout of frmGlyph.  Possibly use multiple forms. (5). Sort of Done.
' Dynamic Creation (where user chooses what Hidden Layers/Neurons). (2)
' Jiggling (oo-er) of the net. (9). Done
' Training Setup.  Test Setup. (2)
' Change GetOuput to be better worked (9). Done.
' New testing program.  I propose training on numbers 0-9 on various fonts,
'   and using more neurons in the hidden layers to cope with this increase.
'   Nothing drastic, perhaps a threefold increase. (8). Done
' Remove dead code (1). Done
' ------
' Jiggling done.  GetOutput reworked.  Tested Ok.
' Adapting for training on numbers: Create,Train,Test Battery,Load,Save,Destroy, Test, Output
' Test 9: 651*96*144*10.  25 Training Fonts.  Fitted anti-aliased characters.
' Glyphs trained: 2404.  Save Num1.nnt
' Trained Fonts
' --------------------
' Successes: 179
' Failures: 71
' Average Correctness: 87.78%
' Average Incorrectness: 16.62%
'
' UnTrained Fonts
' --------------------
' Successes: 29
' Failures: 41
' Average Correctness: 65.77%
' Average Incorrectness: 34.35%
'
' Seems to be getting 0,5,6 confused.  Additionally 3,5 are being confused.  Seems to be
' basing recognition of "1" on character width rather than what it looks like.
' Doesn't seem to be getting better with additional training.  Time for some research...
' Clean-up of code done.
' Draw & Test code done.
' Try: Compute Sum of Square Error, stop using Learning Coefficient and use adaptive
' learning instead.  Hum.  Before I do this (don't understand it at the mo), I'll try using a lower
' learning co-efficient.  Fed up with random font selection in training.  Order instead.
' Test 10: 651*96*144*10.  25 Training Fonts.  Fitted anti-aliased characters. Learning Coeff = 1.4
' Glyphs trained: 11000 (overnight run).  Save: wt1.nnt
'Trained Fonts
'--------------------
'Successes: 249
'Failures: 1
'Average Correctness: 99.43%
'Average Incorrectness: 0.27%
'
'UnTrained Fonts
'--------------------
'Successes: 53
'Failures: 17
'Average Correctness: 79.75%
'Average Incorrectness: 16.08%
'
' Plus, drawn glyphs have good accuracy too.  Woowoo.
' Possibly my tests for Success/Failure are a little too stringent on untrained.
' Who cares, I'm happy with this.
' Well...ish
' Let's use all the fonts I've got that I can read at small pitch, whether or not
' they're italic, handwriting font, typewriter font.  Whatever.
' I think I'll change battery tests tolerance to 70% & 30% too.
'--------
' Major change in neural net code.
' Start using arrays for massive speed increase!!
' Ok, this is done, but the code is nasty and complicated (as you might expect).
' Timing:
'  Create: 2566ms, Output: 245 ms, Train: 950ms, Destroy: 13ms, Load: 2051ms, Jitter: 500ms, Save: 2087ms
' Ulli output request done (compiler constant).
' Now using Combo1.Sorted property instead of doing it myself (removal of modArraySorts.bas)
' New net code training tests.  Well it seems to work ish.
' Learning of zero tending to zero!  Doesn't seem to happen on re-train.
' Introduced another jiggling function - KickZeros.  Doesn't work.
' New randomiser (Nguyen-Widrow).  AMAZING!  SO MUCH BETTER!!! (Who'd have thought it eh?).
' Ok, let's set up a test:
' Test 10: 651*96*144*10. 51 Training fonts. Fitted antialiased characters.
' AnnealEpoch = 2 Training Epochs, AnnealInc=1.05,AnnealDec=0.97, Initial LC=0.9. Nguyen-Widrow init.
' Glyphs trained: 3063
' Save: wt1.nnt ; Save battery test: Test10.txt
'Trained Fonts
'--------------------
'Successes: 465
'Failures: 45
'Average Correctness: 91.21%
'Average Incorrectness: 5.91%
'
'UnTrained Fonts
'--------------------
'Successes: 54
'Failures: 6
'Average Correctness: 89.12%
'Average Incorrectness: 9.15%
'
' I assume improved success is due to Nguyen-Widrow initialisation rather than
' anything else I've done.  This really needs to be trained on a _lot_ more glyphs.
