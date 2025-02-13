# Visual Basic Guide <!-- omit from toc -->
## A GCSE Visual Basic Guide with many useful tips that I have picked up along my time programming <!-- omit from toc -->

So, you're here because you don't know Visual Basic very well, but you want to. I appreciate that! VB is widely considered a _kinda useless_ language, but for some reason it is one of the options on the GCSE AQA Computer Science specification. And you know what? It sort of makes sense. It is a great starting point for anyone looking to code, and it gives you a great idea of the .NET framework so you can eventually move on to C#, arguably the most useful programming language. So, for now, let's think of it as a stepping stone to a better time, and I hope this guide helps you to understand VB a bit more than you did before.
Thanks for visiting this guide. If you see anything wrong, think anything is missing, or just generally have any other suggestions, please do not hesitate to [open an issue](../../issues) and let me know. With that said, enjoy the guide!

I have attached links to specific chapters below, just click on any link to see more about a topic.

> [!IMPORTANT]
> Please open an issue in this repo if you think I should add something else to this guide. I aim to try to maintain this for as long as possible!

## Contents <!-- omit from toc -->
- [Subroutines](#subroutines)
  - [Defining A Subroutine](#defining-a-subroutine)
  - [Calling A Subroutine](#calling-a-subroutine)
  - [Parameters](#parameters)
- [Functions](#functions)
  - [Defining A Function](#defining-a-function)
  - [Calling A Function](#calling-a-function)
  - [Function Parameters](#function-parameters)
  - [Return Values](#return-values)
- [Structures / Records](#structures--records)
  - [Defining Structures](#defining-structures)
- [Variable Iteration](#variable-iteration)
- [Selecting Cases](#selecting-cases)
- [Modulo (MOD) Operator](#modulo-mod-operator)
- [Uses Of Data Types](#uses-of-data-types)
- [Naming Conventions](#naming-conventions)
- [How To Name Your Variables](#how-to-name-your-variables)
- [Colouring text \& backgrounds](#colouring-text--backgrounds)
- [Random Number Generation](#random-number-generation)
- [Concatenation](#concatenation)
- [Arrays \& Lists](#arrays--lists)

## Guide <!-- omit from toc -->

### Subroutines
Subroutines help massively with organising your code. I have seen code before with a `Sub Main` that is over 300 lines long and it *physically hurts* when people tell me they don't know how subroutines work. They are incredibly simple and people seem to overcomplicate them a lot.


#### Defining A Subroutine
To define a subroutine, simply go outside of `Sub Main` but inside `Module Program`, and type in the following code:

```vbnet
Module Program

    Sub Main()
    
    End Sub

    Sub MyNewSub()
        'Do Something
    End Sub
End Module
```
Personally, I like to put my custom subroutines below Sub Main, just to give me an idea of the order of execution, but they can go anywhere within the module.

#### Calling A Subroutine
Take the sub we made in the previous step, `MyNewSub()`. This can be called **anywhere** in your code (but usually in `Sub Main` if you are making a simple program), simply by typing its name and the brackets, as so:
```vbnet
Sub Main()
        MyNewSub()
    End Sub

    Sub MyNewSub()
        'Do Something
    End Sub
```
This would run `MyNewSub()` whenever `Sub Main` runs, so in practice that would be when the program starts

#### Parameters
Parameters are essentially values that you 'pass over' to a subroutine. For example, if I had an integer variable `x` in one of my subroutines:
```vbnet
    Sub Main()
        Dim x As String = 10
        MyNewSub()
    End Sub
```
and I wanted to use my new subroutine to, say, multiply that value by 2. But there's a problem: variables created within subroutines are stuck inside the sub they are in. Think of it as if the `Sub` and `End Sub` blocks act as a 'wall' around the variable; it can't escape. Unless, of course, you give it a hole that is the perfect shape & size for it to fit through. This is effectively what parameters are.
So, to pass over my variable `x` to my new subroutine, I would do this:

```vbnet
    Sub Main()
        MyNewSub(x)
    End Sub

    Sub MyNewSub()
        'Do Something
    End Sub
```

But there's another problem here: **my new sub is not expecting a value**. So, this is where the second part of parameters come in. What I'd do in this scenario is the following:

```vbnet
    Sub Main()
        MyNewSub(x)
    End Sub

    Sub MyNewSub(x As Integer)
        x *= 2 'See the "variable iteration" chapter of this guide for an explaination of this line
    End Sub
```

> [!NOTE]
> I specify the data type in the part where `MyNewSub` expects a variable. This is important because if you don't do this, VB doesn't know if it will have a string, integer, boolean, date, time, or something else, and it doesn't like that.

> [!NOTE]
> Also notice how when I am calling the subroutine, I don't need to specify the data type. This, again, is important.

I can then use this data **inside the subroutine** to output or calculate more things. However, I cannot normally pass the multiplied value back to the original subroutine using regular subroutines. To do that, I would have to use a [function](#functions)

### Functions
Functions are just [subroutines](#subroutines) where a value is returned to the place where the function was called from. In most programming languages, subroutines are just functions that don't return a value, but for some reason the VB developers felt the need to separate the two. So, we have to work with that.


#### Defining A Function
To define a function, it is the same as [defining a subroutine](#defining-a-subroutine), but you write 'Function' rather than 'Sub'. Easy, right? Just like this:

```vbnet
Module Program

    Sub Main()
    
    End Sub

    Function MyNewFunction()
        'Do Something
    End Function
End Module
```

#### Calling A Function
Again, this is done the same way as [calling a subroutine](#calling-a-subroutine), just by writing its name with the brackets (space for parameters):
```vbnet
    Sub Main()
        MyNewFunction()
    End Sub

    Function MyNewFunction()
        'Do Something
    End Function
```

#### Function Parameters
These behave exactly the same as subroutines, with absolutely no differences. If you are unsure of how those work, have another read [here](#parameters).

#### Return Values
This is where functions get interesting. In a subroutine, the parameters are stuck in the sub once execution has finished. So if I multiply a variable by 2 in a sub, the original parameter still keeps its original value in the origin subroutine.
In a function, though, a return value is **required**. This means that functions can be used for variable assignment, like so:

```vbnet
    Sub Main()
        Dim x As Integer
        x = MyNewFunction()
    End Sub

    Function MyNewFunction()
        'Do Something
        Return something
    End Function
```

But, there's an issue here: I have told the program that my variable `x` will be an integer, but the function could *technically* give another value as its return value. So, we should specify what data type the function will output, like this:

```vbnet
    Function MyNewFunction() As Integer 'Notice that this line specifies the data type now
        'Do Something
        Return something
    End Function
```

So yeah, that's functions. They're really not too complicated. You should use them more. They are very helpful and keep your code more organised. Plus, you need to know them for your GCSEs. So it's a win-win situation really.

### Structures / Records
Structures & records both refer to the same thing. I will call them structures for the sake of this guide, since that's what they're called in VB. But you may hear either term used to describe them.

I like to think of structures as adding 'attributes' to data. For example, if I am trying to create a program that stores data about people:

```vbnet
Console.WriteLine("Enter name")
Dim name As String = Console.ReadLine()
Console.WriteLine("Enter age")
Dim age As Integer = Console.ReadLine()
```
This would work fine, for one person. But for every new user, I need a new set of variables. Not great.

Maybe I could use lists or arrays?
```vbnet
Dim peopleNames as New List(Of String)
Dim peopleAges as New List(Of Integer)

Console.WriteLine("Enter name")
peopleNames.Add(Console.ReadLine())
Console.WriteLine("Enter age")
peopleAges.Add(Console.ReadLine())
```
This also works, but it still isn't great. I have multiple lists! That's so annoying to reference every time I need to use the data.

So, this is a perfect place for a structure.

#### Defining Structures
Defining a structure is very similar to [defining a subroutine](#defining-a-subroutine), but not quite the same. They must be within a module, but not within a subroutine.
 Let's make a structure to contain the data I was trying to get in the previous step:

```vbnet
Structure UserData
    Dim name as String
    Dim age As Integer
End Structure

Sub Main()
    Dim person As New UserData 'Make sure you define it as 'new' then the name of your structure
    Console.WriteLine("Enter Name")
    person.name = Console.ReadLine() 'The word after the dot refers to a variable inside your structure
    Console.WriteLine("Enter Age")
    person.age = Console.ReadLine()
End Sub
```
So now I have one variable, `person`, storing both the name and age of a given person. I can then add this person to a list of people, like so:

```vbnet
Sub Main()
    'Previous code
    Dim people As New List(Of UserData) 'Define the list as a new list of your structure
    people.Add(person)
End Sub
```

Then if I want to reset the person variable to reuse it, I can, with the following line:
```vbnet
Sub Main()
    'Previous code
    person = Nothing
End Sub
```

And that's structures folks! Again, relatively simple, but very useful.


### Variable Iteration

To increase the value of a variable `x` by 1, you may have been taught the following: 

```vbnet
x = x + 1
```

But, instead of this, you can do:
```vbnet
x += 1
```

This works for other operations too:
```vbnet
x -= 1 'subtracts 1 from x
x *= 1 'multiplies x by 1
x /= 1 'divides x by 1
```
Of course, some of these are completely useless (such as dividing by 1), but you can replace 1 with any other value, including another variable:
```vbnet
x += y
```

### Selecting Cases
`If` statements are incredibly powerful. However, they can be time consuming to type out, and awkward to use, with long statements. So, in some instances, a `Select Case` block is more suited for the job. Take this `If` statement:
```vbnet
If input = 0 Then
    Do something
Else If input = 1 Then
    Do something else
Else If input = 2 Then
    Do another thing
Else
    Do something if none of the above were true
End If
```

This could be made much neater using a select case block:
```vbnet
Select Case input
    Case 0
        Do something
    Case 1
        Do something else
    Case 2
        Do another thing
    Case Else
        Do something if none of the above were true
End Select
```
> You can also write lines such as `Case > 3`, which does exactly what you'd expect

Both of these code blocks do the **exact same thing** as each other, which makes coding things such as user inputs, where you may want a user to input, say, 0, 1, 2, 3, or "exit" to select an option from a list. It just helps to make code look a bit neater.

However, if statements can be more powerful in certain scenarios, so you should figure out which works best for your use case

### Modulo (MOD) Operator
This is a very helpful arithmetic operator that seems to confuse many people, despite its simplicity. When you do a division manually, using the bus stop method, you are left with a remainder if the number does not divide evenly. This remainder is the value outputted by the MOD operator.
Say I have two variables, x and y. If I want to get the remainder when x is divided by y, I would do this:
```vbnet
Dim x As Integer = 11
Dim y As Integer = 5
Console.WriteLine(x Mod y) 'This would output 1, since 11 / 5 = 2, remainder 1
```

### Uses Of Data Types
Most coding languages contain the same data types, since they are key concepts of basic computer science. Below is a list of a few different common data types that you may find useful:

| **Data Type** | **VB** | **What It Stores**                                                            |
|---------------|------------|---------------------------------------------------------------|
| Boolean       | Boolean    | True or False                                                                 |
| Character     | Char       | Single character                                                              |
| Date & Time   | Date       | Date & Time from midnight on January 1, 0001 to 11:59:59 on December 31, 9999|
| Decimal       | Decimal    | Real numbers                                                                  |
| String        | String     | A string of characters                                                        |

There are more than just these, but these are the most commonly used in GCSE-level programming. The rest of the data types can be found in the Visual Basic docs

### Naming Conventions
This isn't so much a VB tip as it is a general coding tip. Naming your variables, subroutines, functions, and anything else you need to name whilst coding is very important. So, most coders use a similar system for variable naming.

For single-word variables, you should write the name all lowercase, for example:
```vbnet
    Dim word As String
```
For multi-word variables, you should write the first word lowercase, and the rest of the words with the first letter capitalised, like so:
```vbnet
    Dim thisIsALongVariable As Char
```
Or
```vbnet
    Dim wowThisIsVeryLongAndCouldProbablyBeShortenedSignificantlyButIWantedToMakeItAsComicallyLongAsPossible As Date
```

And for functions, subroutines and structures, you should have all of the first letters capitalised, regardless of the number of words:
```vbnet
    Function ThisIsAFunction() As Integer
    End Function

    Sub ThisIsASub()
    End Sub

    Structure WowLookAStructure
    End Structure
```
These aren't really as important as some of the other tips, but they help to make your code more consistent & predictable, along with being easier for other coders to understand.

### How To Name Your Variables
This is just a quick little tip, but make sure to name your variables with sensible names. In computing, we call these 'meaningful identifiers', which ensure that you know what a variable is for. For example, instead of doing this:
```vbnet
    Dim myVariable As Integer
```

You could do
```vbnet
  Dim lengthOfSide As Integer
```

These are just little things that make debugging one heck of a lot easier, and make it easier for others to read.

### Colouring text & backgrounds
If you want to change the colour of the text in a VB console window, you reference the `Console.ForegroundColor` property. You can either set this to an integer between 0 and 15, with each integer corresponding to a colour as listed below, or by using a ConsoleColor.example, replacing 'example' there with a colour from the list below. There are 16 predefined colours, outlined in this table:

| **ConsoleColor**        | **Integer Value** | **Description**                                      |
|:-----------:|:-----:|--------------------------------------------------|
| Black       | 0     | The color black.                                 |
| DarkBlue    | 1     | The color dark blue.                             |
| DarkGreen   | 2     | The color dark green.                            |
| DarkCyan    | 3     | The color dark cyan (dark blue-green).           |
| DarkRed     | 4     | The color dark red.                              |
| DarkMagenta | 5     | The color dark magenta (dark purplish-red).      |
| DarkYellow  | 6     | The color dark yellow (ochre).                   |
| Gray        | 7     | The color gray.                                  |
| DarkGray    | 8     | The color dark gray.                             |
| Blue        | 9     | The color blue.                                  |
| Green       | 10    | The color green.                                 |
| Cyan        | 11    | The color cyan (blue-green).                     |
| Red         | 12    | The color red.                                   |
| Magenta     | 13    | The color magenta (purplish-red).                |
| Yellow      | 14    | The color yellow.                                |
| White       | 15    | The color white.                                 |

So, for example:

```vbnet
Console.ForegroundColor = ConsoleColor.DarkGreen
Console.ForegroundColor = 2 'both of these lines do the same thing: set the text to dark green
```

And to set the background colour, you reference the `Console.BackgroundColor` property in the same way, for example:

```vbnet
Console.BackgroundColor = ConsoleColor.Magenta
Console.BackgroundColour = 13 'again, both lines set the background colour to magenta
```

> [!NOTE]
> The background colour doesn't affect the whole console window unless you try to make it. It only affects the lines after you set the colour, and only when something is outputted to the console. To see what I mean, give it a try in a VB console app.

### Random Number Generation
There are a few ways to go about generating random numbers in Visual Basic, but my preferred method (and, _technically_, the most random method, although unless you are doing advanced cryptography or need INCREDIBLY random numbers this doesn't really matter) is the one I will outline. Say I have an integer variable, `x`, that I want to be random number in between two other variables, `min` and `max`:
```vbnet
Dim x, min, max As Integer
min = 1 'The minimum value
max = 10 'The maximum value

Randomize()

x = Int(Rnd() * max) + min
```
To generate a random number using this method, the `Randomize()` must be called only once at any point before generating the number. I like to put it right at the top of my `Sub Main` in most programs I create
This code would assign a random integer value that is greater than or equal to the value of `min` (1, in this instance), and less than but **not equal to** the value of `max` (10 in this instance). Or, in [interval notation](https://en.wikipedia.org/wiki/Interval_(mathematics)#Notations_for_intervals), [1,10)

### Concatenation
Concatenation is the joining of multiple strings together, and in VB is relatively simple, and the language gives you a few options of how to do it.
For example, if I wanted to join the string "Hello" with the string "World", with a space in the middle, I could do any of the following:

```vbnet
Console.WriteLine("Hello" + " " + "World")
Console.WriteLine("Hello" & " " & "World")
Console.WriteLine("{0} {1}", "Hello", "World")
```
But what if I have two variables, or one predefined string and one variable? Then, I would do this:
```vbnet
Dim x As String = "Hello"
Dim y As String = "World"

Console.WriteLine(x & " " & y)
Console.WriteLine(x + " " + y)
Console.WriteLine("{0} {1}", x, y)
Console.WriteLine($"{x} {y}") 'notice the $ symbol after the bracket but before the double quotes
```

### Arrays & Lists
Arrays are just a collection of the same data type, like a list of values. Lists are just variable-size arrays.
When defining an array, you must specify the size of the array. This cannot be changed later on, except with some janky techniques which you should never do.
To define an array in VB, you can simply type:
```vbnet
Dim myArray(size) As datatype
```
Replace "datatype" with any data type or even the name of a [structure](#structures--records)
Bear in mind that the size of the array is zero-based, so an array of size 0 can store 1 item, and an array of size 1 can store 2.
Then, if I wanted to have an array of, for example, strings, I could do the following:
```vbnet
Dim colours(6) As String
colours(0) = "Red"
colours(1) = "Orange"
colours(2) = "Yellow"
colours(3) = "Green"
colours(4) = "Blue"
colours(5) = "Indigo"
colours(6) = "Violet"
```
I could then reference each position in the array as if it were just a variable:
```vbnet
If colours(0) = "Red" Then
    Console.WriteLine("This is good...")
End If
```
---
_✨ awaiting more content ✨_
