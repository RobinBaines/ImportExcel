OOO3 is an application which simulates some aspects of Object-oriented Ontology, which is described in the book Object-oriented Ontology by Graham Harman.
Making a program based on OOO involved making decisions on implementation which may not do justice to the original ideas.
However getting to a working program and some of the conclusions based on using it was interesting.

Object Oriented Ontology (OOO) describes Real Objects (ROs), Real Qualities (RQs), Sensual Objects (SOs) and Sensual Qualities (SQs).
The following should clarify how I view the difference between these:

In his book, Object-oriented Ontology, Graham Harman includes a section on an OOO analysis of the American Civil War (ACW).
He uses the ACW as an example of an RO. I assume here that he wrote the chapter based on his SO(s) of this war.
I further assume that he created an SO, I call it an Event, which records his experience while writing the chapter.
On completion, the section in the book is an RO. The RO includes well established facts, dates and figures but also an OOO analysis of the ACW 
illustrating how looking for 'symbiosis' and the lifecycle of an object can result in interesting insights.

While reading about the ACW in the book I create an Event SO, let's call it ReadingACW.
By chance I read a book about the ACW last year. I therefore had an existing ACW SO.
The information in the chapter augmented this SO with SQs and made references between the ReadingACW SO
and my existing ACW SO, recording not only the SQs as stand-alone statements, but also where they came from. 

In the software and what follows I use Event with a capital letter when referring to an SO and event when referring to the sub-events in an Event.
The 'Event' SO is like a theatre with sub-events describing changes to SOs and interactions between them.

This software simulates creation of SOs and SQs using a stream of data in a script file.
It uses object oriented software ideas such as inheritance, overriding and properties.
I have chosen to illustrate this by describing several Events starring 2 dogs we own and local places for dogs to roam and play.
These Events are in a script file called script.txt. It is possible to describe other Events in other script files. 
ACW.txt describes the American Civil War and is included in the release.

In my experience the stream of data in a (subject-object) script may have different sources: spoken language, written language, thoughts, actions,
and emotion and external events and processes involving visual stimulus, smells, sound.... 
In a subject-object Event I may switch rapidly between different sources 
but I seem to process data in a sequence as opposed to, for example, multiple parallel streams.  
My stream depends on my interpretation of ROs and on my 'state' or in this context, on the state of my SOs including my emotional state. 
One day I will be receptive to the external stimuli the next I may be thinking about something and the external world makes less of an impression.

I have not spent time on using sophisticated streams of data from different sources; besides being very difficult/impossible it is not the focus here.
The stream in a script must conform to a simple syntax which is explained in the file Program.cs.
The example script in the file called script.txt, is annotated and is more than enough information to allow writing of other scripts 
which conform to the syntax.
The script as a 'reality' simulator is an extreme simplification and it may be such a simplification that the idea of this program 
has no point at all. On the other hand some of the results and conclusions look interesting; see below.

Use by entering 

        OOO3.exe Script.txt > ScriptOutput.txt 
		
on the command line.

Or if using the Scripts folder

		ooo3.exe ..\Scripts\script.txt > ..\Scripts\Output\ScriptOutput.txt

The program interprets the script.txt and writes the output to the console or to a file using redirection.
The output is managed using output commands at the start of the script:

	//Output commands.
	DisplaySOs *	//display all the SOs with SQs.
	DisplaySOs Dog  //display the SOs with SQs which inherit from Dog.

	QuerySOs Event, Dog  //display all the SOs which inherit from Event and which include an SO which inherits from Dog.
	
	GenerateScript Event //Look for SQ's in SOs derived from Event which are an SO. For each of these generate a script to add an SQ 
						 //to the SO. For example 
							MeetBakkeveen : Event {
								Nida = playing Siena
								///will create 
								Nida{
									playing = Siena
							The idea is that if Nida was playing with Siena then Nida has the SQ (can) play.
	
	RandomWalk 30	//Print a list of SQs by starting with a random SQ then moving to an SQ with the same Name but another SO.
					//Do this 30 times.
					
	WOOOF.exe has a windows forms user interface. It uses OOO3.exe to process a script, creating the output files while animating the process. 

SOME OOO IDEAS:

A flat ontology implies that interaction between inanimate objects without an observer is valid. 
The example script implies Events with myself as observer.
There is no reason not to create a script with combinations of animate and inanimate objects with no observer. For example a script describing a Dogs 
interaction with food and may be one with 2 or more inanimate objects with no observer. Or even quantum events?
See Water.txt and WaterDetail.txt.

Vicarious causation: OOO requires an SO to facilitate interaction between (inanimate) objects. In my interpretation this occurs in the Event SO.
Writing a script with only inanimate object immediately raises questions about the nature of the Event. Does it have qualities: a Place, datetime...?

Undermining means that (real) objects can't be reduced to only their parts. 
Nida the dog in the script is only sketched out with a limited number of SQs. I could go on describing or remembering her 
but I will never get to what it is like to be a dog called Nida.

Overmining means that an object is less than it's effects. 
 
Two ideas from OOO analysis which play a role in the ACW script. 
Symbiosis. 2 or a single event which causes an irreversible change in a real object.
Object lifecycle. Irreversible changes in the state of an object often marked by symbiosis.

SOME SOFTWARE IDEAS:

1. Inheritance allows an object may be derived from another, often, more abstract object. 
For example Nida inherits from Dog, Me inherits from Person and Dog and Person both inherit from Animal.
Inheritance can be thought of as a 'type of' relationship: Nida is a type of Dog and Dog is a type of Animal. 
The convention 'Nida : Dog'  and 'Dog : Animal' is used to represent inheritance in a script.

A predefined class hierarchy is used for the class Sensual Object.

    SensualObject inherits from BaseClass and
    
    SensualQuality also inherits from BaseClass
    
BaseClass encapsulates properties which are part of the SensualObject and the SensualQuality. 

In the Object Oriented Software Language called c#, classes are predefined and objects are instantiated from the classes when the software runs.
In this software all SOs are instantiated from the predefined SensualObject class but inheritance from one SO to another is 
simulated using a reference in the BaseClass.

For example Nida is an SO. Nida is a Dog, and Dog is an SO. Dog is an Animal and Animal is an SO.
These inheritance relationships are captured using a reference from Nida to Dog and from Dog to Animal.

2. Overriding in c# means replacement or enhancement of a function or property value in a derived object.
Here overriding is supported by allowing an SQ defined in an SO to be overridden in a derived SO. 
For example a Dog may have an SQ called Bark = True, implying that (nearly) all dogs bark.
Nida who is derived from Dog may have SQ Bark = Loud; and may also have another SQ: Bark = High.
A dog which cannot/does not bark would have SQ Bark = False.

3. References: An SO or SQ may refer to a number of SOs.
For example an Event called MeetBakkeveen (MeetBakkeveen : Event) may refer to
the SO Nida and SO Siena. After processing of the script these SOs are coupled to the Event using references in the MeetBakkeveen SO.
*Bakkeveen is a place in Friesland.

The following captures an observation about an SO which occured during the Event.
    Nida = swimming //is a quality of Nida which occurred during the MeetBakkeveen : Event.
	
Slightly different is that I may realise that an SO has an SQ during an Event:	
		
  	Nida {
  		FrontLeftLeg	

The SQ is added to Nida and an SQ is added to the Event to record that happening (see redundancy below).

Sometimes an SQ may refer to an SO. For example

	MeetBakkeveen: Event {
        	Place = Bakkeveen

Bakkeveen is a pre-existing SO defined as Bakkeveen : Place
In this case the reference is stored in the SQ of the MeetBakkeveen Event.

4. Names: An SO and an SQ are identified by a unique name or phrase (a string which may include white space). 
There is no attempt to enumerate possible values of SQs. For example there is no list of possible values for the Dog.Breed.
However it is possible to query the SOs for all the values used in SQ Breed.

5. Timestamping: The BaseClass has a Time stamp property. By default this is set to the computer DateTime when the SO or SQ 
is created. It may also be overridden to create a more realistic time sequence.
The time stamp looks like this in the script:

    Time 2025-02-04 16:30:00.000

This time will be recorded with every following state change of SOs or SQs until the following Time value occurs in the script.

OOO proposes a flat ontology and interaction between inanimate objects is valid. 
The Event object used in this program could be used to coordinate this sort of interaction.
While distinct events may occur in these types of Event, processes are also common.
The script language does not support descriptions of processes except approximately by using time stamps. For example growth of a crystal
could be described by 

    Crystal = size 5
    
    Time 2000-02-04 06:30:00.000
    
    Crystal = size 6
    
    Time 2025-02-04 20:30:00.000

6. No forward referencing. The program processes the script step by step without forward references.
For example: 

//The class Event is defined with several SQ's.

	Event {
	Place
	Weather
	End
	}

//Only then can Init be defined as a type of Event.

	Init : Event {			//Init is a type of Event


INITIAL CONCLUSIONS
One idea of the experiment is to look at the data structure, sub-routines and functions needed to turn the list of input data into 
objects and to assess whether this overhead is reasonable. The BuildSOs class processes input data and builds the SOs in a List array 
called 'theSOs' in a predefined class called SOs. The BuildSOs and SOs classes can be thought of as 'Me' as it includes all the SOs and
the processing required to manage them. What to call this when a script contains only inanimate SOs?

One problem with the result is that the SOs and SQs are not 'visible' and can only be inferred from the output (sounds familiar!).
For this reason the SOs class also includes several subroutines which query the object model. 
These subroutines use similar approaches by iterating through SOs and SQs.

    DisplaySOs(string Parent) Display the SOs with SQs filtering on Parent. For example SOs.PrintSOs("*"); displays all the SOs. 
    SOs.PrintSOs("Dog"); Displays SOs derived from Dog.
	Here is the SO Nida at the end of the script (Nida undergoes several changes during the script; check out the time stamps.
	
	//////////////////////////////////////Nida 21 //////////////////////////////////////
	2025-02-02 14:00:00.000 Nida : Dog
	SO Nida has qualities.
		2025-02-02 14:00:00.000 MeetNida => Tail = Short
		                        MeetNida => Colour = Black
			2025-01-01 16:00:00.000 SO Reference Black : Colour
		                        MeetNida => Name = Nida
			2025-02-02 14:00:00.000 SO Reference Nida : Dog
		                        MeetNida => Name = Nidaatje
		2025-02-04 14:00:00.000 DogsAreDogs => INHERIT_SENSUALOBJECT Dog = True
		                        DogsAreDogs => Breed = None
		                        DogsAreDogs => Size = Medium
		2025-02-04 16:00:00.000 MeetBakkeveen => Colour = Black
			2025-01-01 16:00:00.000 SO Reference Black : Colour
		                        MeetBakkeveen => faster = Siena
			2025-02-03 14:00:00.000 SO Reference Siena : Dog
		2025-03-23 10:43:51.955 Generated => eating = True
		                        Generated => faster = Siena
			2025-02-03 14:00:00.000 SO Reference Siena : Dog
		                        Generated => playing = Siena
			2025-02-03 14:00:00.000 SO Reference Siena : Dog
		                        Generated => sleeping = True
		                        Generated => playing = Black Dog
			2025-02-04 16:10:00.000 SO Reference Black Dog : Dog
		                        Generated => coming = True
		                        Generated => sitting = True
		                        Generated => swimming = True
	
	SO Nida has references.
		2025-01-01 16:00:00.000 SO Reference Colour
		                        SO Reference Black : Colour
		2025-02-02 14:00:00.000 SO Reference MeetNida : Event
		2025-02-03 14:00:00.000 SO Reference Siena : Dog
		2025-02-04 14:00:00.000 SO Reference DogsAreDogs : Event
		2025-02-04 15:00:00.000 SO Reference MeetBakkeveen : Event
		2025-02-04 16:10:00.000 SO Reference Black Dog : Dog
		2025-02-04 17:40:00.000 SO Reference faster
		2025-03-23 10:43:51.955 SO Reference Generated : Event
		                        SO Reference playing : Verb
	
	    
    QuerySOs(string _parent, string _child) Query SOs for example SOs.QuerySOs("Event", "Dog"); looks for SOs with a Parent = "Event" and referenced SOs 
	    with a Parent = "Dog".
		Query Parent: 'Event' for Child: 'Dog'
			 Parent: 2025-02-02 14:00:00.000 MeetNida Child: 2025-02-02 14:00:00.000 Nida
			 Parent: 2025-02-03 14:00:00.000 MeetSiena Child: 2025-02-03 14:00:00.000 Siena
			 Parent: 2025-02-03 14:00:00.000 MeetSiena Child: 2025-02-02 14:00:00.000 Nida
			 Parent: 2025-02-04 14:00:00.000 DogsAreDogs Child: 2025-02-04 14:00:00.000 Dog
			 Parent: 2025-02-04 14:00:00.000 DogsAreDogs Child: 2025-02-02 14:00:00.000 Nida
			 Parent: 2025-02-04 14:00:00.000 DogsAreDogs Child: 2025-02-03 14:00:00.000 Siena
			 Parent: 2025-02-04 15:00:00.000 MeetBakkeveen Child: 2025-02-03 14:00:00.000 Siena
			 Parent: 2025-02-04 15:00:00.000 MeetBakkeveen Child: 2025-02-02 14:00:00.000 Nida
			 Parent: 2025-02-04 15:00:00.000 MeetBakkeveen Child: 2025-02-04 16:10:00.000 Black Dog
			 Parent: 2025-03-05 09:48:48.058 Generated Child: 2025-02-02 14:00:00.000 Nida
			 Parent: 2025-03-05 09:48:48.058 Generated Child: 2025-02-03 14:00:00.000 Siena
	    
    RandomSQs() Display a list of SQs by starting with a random SQ then moving to an SQ with the same Name but another SO which has not been printed yet. 
    When there are no SQs meeting the above criterion switch to another SO which is referenced by the last SO.
	This results in a sort of declamation of related qualities: 
	SOs RandomSQs 
	Using 235 qualities and starting with index = 198
		SO Feel with SQ Root = Feel
		SO Verb with SQ Root = True
		SO playing with SQ Root = play
		SO sleeping with SQ Root = sleep
		SO running with SQ Root = runn
		SO call with SQ Root = call
		SO come with SQ Root = come
		SO coming with SQ Root = com
		SO sit with SQ Root = sit
		SO sitting with SQ Root = sitt
		SO eating with SQ Root = eat
		SO swimming with SQ Root = swimm
	
	SWITCH from swimming.Root using reference Generated to Me.Feel
		SO Me with SQ Feel = Cold
	
	SWITCH from Me.Feel using reference InitEvent to Me.Gender
		SO Me with SQ Gender = Male
		SO Animal with SQ Gender = True
		SO M with SQ Gender = Female
	
	SWITCH from M.Gender using reference MeetNida to MeetNida.Weather
		SO MeetNida with SQ Weather = Cold
		SO Event with SQ Weather = True
		SO MeetSiena with SQ Weather = Cold
		SO MeetBakkeveen with SQ Weather = Cold
	
	SWITCH from MeetBakkeveen.Weather using reference Black Dog to Black Dog.Size
		SO Black Dog with SQ Size = Small
		SO Nida with SQ Size = Medium
		SO Siena with SQ Size = Large
		SO Dog with SQ Size = True
	
	SWITCH from Dog.Size using reference DogsAreDogs to DogsAreDogs.Place
		SO DogsAreDogs with SQ Place = Home
		SO Event with SQ Place = True
		SO MeetNida with SQ Place = Bakkeveen
		SO MeetSiena with SQ Place = Dino Bos
		SO MeetBakkeveen with SQ Place = Bakkeveen
		SO T2 with SQ Place = Home
	
	SWITCH from T2.Place using reference Verb to Verb.Root
		SO Verb with SQ Root = True
	

Inheritance: I remember Fons commenting that some people with autism are not able to generalise; which could mean that every dog is a new
SO. 
An object is a collection of other objects. This is problematic in a practical solution. 
The WaterDetail script describes collisions of water drops. Defining a collection of waterdrops seems to be necessary.
Should the script describe the parts of the drop by tracing the events which resulted in its creation or should it be a list 
of its parts? 
Inheritance offers a short cut. It is possible to define an object WaterDrop and then define WD1 : WaterDrop, WD2 : WaterDrop, WD3 : WaterDrop etc 
as instances of waterdrops which play a role in the script. Any detail of drop formation or water vapour creation from water in the sea and ultimately
water creation from Hydrogen and Oxygen can be hidden or made explicit in the WaterDrop object.

Juliet's immediate response to OOO3 was that all the output could be generated directly from the script.
This must be True but is it a problem?

1. This could also be true of humans. The hidden nature of SOs and the number of SOs in our brain coupled with memory loss makes this difficult/impossible to judge.

2. A 'Me' SO could include emotions which influence creation of other SOs. No emotion no memory!? 
There is no attempt to simulate this in the software.

3. An Init event is used to populate the 'Me' SOs before processing of other Events takes place. Altering the Init SO will influence creation of other SOs.

4. The object structure is easier to use by the output routines such as DisplaySOs, QuerySOs and RandomSQs, than event lists in a script. 
The object model includes redundancy. For example Nida.Tail = Short occurs in Event MeetNida. 
An SQ is included in Nida which couples the SQ in Nida to an SQ in the event MeetNida. 
Referencing MeetNida from Nida and vice-versa is easy and that seems analogous to thinking about Nida brings to mind MeetNida and vice-versa.

5. Complexity and emergence. 
Is turbulence in faster flowing water an emergent property? Computer simulation of turbulent flowing water is only approximate, the Navier Stokes Equations do
describe it accurately but are proving impossible to solve. Perhaps this program highlights the extreme simplifications required 
of any simulation of thought processes. Using this program with 10000 SOs (5 SOs per day * 365 * 50 years of SOs) and 20 SQs per SO would also not help 
and would still be an extreme simplification of what is essentially an analogue process. 
The dependence of our experience on emotion excludes any form of experience in a simulation. The most we can expect is 
surprise at relationships a program can find hidden in the input data.

6. Simulation of Thinking. 
The program and script simulate this in 2 ways by defining generalised objects like Dog on the fly in the script and 
by analysing Events looking for characteristics of other SOs.

In more detail: 
a. Create objects like Dog, on the fly, simulating the realisation that several SOs are the same type: Nida and Siena are a type of Dog.

The Dogs Nida and Siena are defined before the SO Dog has been defined. Here is Nida who is defined in event MeetNida.

	Nida {
  
		FrontLeftLeg
		BackLeftLeg
		FrontRightLeg
		BackRightLeg
		Tail = Short
		Bark
		LeftEar
		RightEar
		Colour = Black
...
		}

In Event DogsAreDogs : Event {

Dog is defined

	Dog : Animal {    //At home I realise that Nida and Siena are a type of Animal called a Dog.
  
		Breed
		Size
		}

And only then is the derivation of Nida from Dog defined:

	Nida : Dog	{	//Because Nida already exists the default qualities (qualities with Value True of not defined) are moved to Dog

		Breed = None
		Size = Medium
		}

The result is the following showing how default properties are now in Dog and the none default properties of Nida remain in Nida.
	SO Nida inherits from Dog which has qualities
 
	//////////////////////////////////////Dog 25 //////////////////////////////////////
	2025-02-04 14:00:00.000 Dog : Animal
	SO Dog has qualities.
		2025-02-04 14:00:00.000 DogsAreDogs => INHERIT_SENSUALOBJECT Animal = True
		                        DogsAreDogs => Breed = True
		                        DogsAreDogs => Size = True
		                        MeetNida => FrontLeftLeg = True
		                        MeetNida => BackLeftLeg = True
		                        MeetNida => FrontRightLeg = True
		                        MeetNida => BackRightLeg = True
		                        MeetNida => Tail = True
		                        MeetNida => Bark = True
		                        MeetNida => LeftEar = True
		                        MeetNida => RightEar = True
		                        MeetNida => Colour = True
		                        MeetNida => Name = True

	SO Dog has references.
		2025-01-01 16:00:00.000 SO Reference Colour
		2025-02-04 14:00:00.000 SO Reference DogsAreDogs : Event

NOTE: There is an SO called Black but how a colour is perceived depends on the object. Nida.Black is different from the SO Black.

b. Analysing Events looking for characteristics of other SOs.
A stream of data is extracted from the SOs and is written to a separate script in an Event called Generated.txt.
The subroutine GenerateAScript() looks for SQs of Events which are SOs and 
creates a script to add the quality to the subject.
For example:

        MeetBakkeveen : Event {
        Nida = swimming
        
		
is used to imply that Nida can swim and will create a new SQ for Nida:

		Nida {
		swimming = True
		}
		
And to imply that swimming is a verb.
		swimming : Verb {
			Root = swimm
		}		

This is an initial attempt at creating SOs to hold types of speech (verbs, adverbs, nouns) and to eventually
add 'production' rules which described how the words should be used in output.
I would need to simulate reading a book about grammar to improve. 

//////////////////////////////////////////
RO Lifetime and the ACW and Water scripts.

The American Civil War and Fort Sumter are both examples of RO which have ended and will not change.
Opinions as to the exact moment this type of object started and ended may differ and are properties/qualities of SOs.
OOO theory seems to depend on comparing an RO with one of many possible SOs. 
Only then can relationships between RO, SO, RQ en SQ be analysed.

The OOO3 program allows (reasonably unlimited) nesting of SOs.
Initially this was used to define an SO of type Event and then to define and then modify SOs which were nested 
in the Event.
For example in ACW.txt the American Civil War Event has a nested Fort Sumter SO (which is also an Event); 
Fort Sumter : Confed_victory : Battle : Event).

ACW : Event {
...
Fort Sumter : Confed_victory {
		CWSAC = Decisive
		...
		}
}

In Script.txt the MeetNida Event defines a nested Nida SO.
MeetNida : Event {
	...
	//define Nida.
	Nida {
			FrontLeftLeg 
	...
	}
}	

The Nida SO is then modified in MeetBakkeveen : Event.
....
MeetBakkeveen : Event {
	...
	//Modify Nida.
	Nida {
		faster = Siena	
		...
		}
}		

With unlimited nesting of SOs it is possible to divide an SO in to other SOs and so on.
In the following, the battle Fort Sumter went on for 2 days and is is possible to divide it into 2 days and these
days in to time periods of 1 hour and those hours into seconds.
But it seems unlikely that there is any sensible way to relate these time period SOs to ROs 
after all what does an RO know about time periods?
Fort Sumter : Confed_victory {
	...
	Day1 : Event {
			Weather = good
			Hour1 : Event {
							Weather = raining
							Minute1 : Event {
								Weather = dryer
								}
							}
		}
	Day2 : Event {
		Weather = sunny
		}		
...
}


The SO Nida the dog may also be seen as an Event. Somewhere at some time she was created, 
for example at the moment of gestation or of birth, but again this highlights how arbitrary these SO qualities
must be when compared to RO qualities; can we have an opinion about when the RO Nida came into being?
However using the SO Nida during an Event 'MeetNida' looks legitimate as Nida was essentially 
the same object during the Event.

Water.txt is a script which describes how 2 drops of water collide to form a single drop which then divides
into 2 other drops some time later when a gust of wind occurs.
	In the 
	
		Collision : Event { 
		
Drop 1 and Drop2 are defined before Big_drop is defined when Drop1 and 2 collided.
In GustOfWind : Event the reverse takes place with Drop3 and Drop4 defined (as part of) Big_drop when that drop split
and ceased to exist.

Here is the object Big_drop.
//////////////////////////////////////Big_drop 9 //////////////////////////////////////
2025-01-01 12:30:00.000 Big_drop : Water
	                       Parent SO = Collision : Event
SO Big_drop has qualities.
	2025-01-01 12:30:00.000 Collision => INHERIT_SENSUALOBJECT Water = True
	                        Collision => mass = 2,3 gm
	2025-01-01 13:00:00.000 GustOfWind => mass = 0 gm

SO Big_drop is part of other objects.
	2025-01-01 12:00:00.000 Is part of Parent = Collision : Event
	2025-01-01 13:00:00.000 Is part of          Drop3 : Water
	                        Is part of          Drop4 : Water
	                        Is part of          GustOfWind : Event
SO Big_drop has references.
	2025-01-01 12:00:00.000 Includes          Drop1 : Water
	                        Includes          Drop2 : Water

And here the object Drop3. 
//////////////////////////////////////Drop3 11 //////////////////////////////////////
2025-01-01 13:00:00.000 Drop3 : Water
	                       Parent SO = GustOfWind : Event
SO Drop3 has qualities.
	2025-01-01 13:00:00.000 GustOfWind => INHERIT_SENSUALOBJECT Water = True
	                        GustOfWind => mass = 1 gm

SO Drop3 is part of other objects.
	2025-01-01 13:00:00.000 Is part of Parent = GustOfWind : Event
SO Drop3 has references.
	2025-01-01 12:30:00.000 Includes          Big_drop : Water


This script does not describe how Drop1 and 2 formed. This is legitimate because we may define the 
initial conditions of any script and in this case the fully formed ROs Drop1 and 2 are initial conditions. 

The script WaterDetail.txt is an attempt to trace the lives of Drops1 and 2 back to the moment they nucleated 
in a cloud raising questions of when these different drop ROs (Drop1, SmallDrop1 and VerySmallDrop1) 
start to exist.

	Drop1 : Water {
	mass = 1 gm
		SmallDrop1 : Water {
			VerySmallDrop1 : Water {
					Molecule1 : WaterMolecule {
						}
					aerosol1 : aerosol {
						}				
					}
			Molecule1_2 : WaterMolecule {
					}	
			Molecule1_3 : WaterMolecule {
					}	
//etc					
			}
	}
	
VerySmallDrop1	is defined as an aerosol (wettable and insoluble) plus a molecule of (condensed) water.
The molecule of (condensed) water is called Molecule1_2 : WaterMolecule
It is 'type of' WaterMolecule which is further defined as being made of the elements H2 and O2.
Here the type of/inheritance is being used to end further reduction of an SO into parts.
It is a way to 'snapshot' the molecule and is a standin for explaining where the molecule came from.
In this example the 'type of' means we do not have to trace the molecule of water to a molecule 
of water vapour which may have come from the sea etc.  
