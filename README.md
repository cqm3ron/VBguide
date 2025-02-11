# Visual Basic Guide <!-- omit from toc -->
## A GCSE Visual Basic Guide with many useful tips that I have picked up along my time programming <!-- omit from toc -->

So, you're here because you don't know Visual Basic very well, but you want to. Or, you're here because I got fed up of you asking me for help and sent you here. I'll let you decide what your reason is, but I hope this guide helps nonetheless.

I have attached links to specific chapters below, just click on any link to see more about a topic.

## Contents <!-- omit from toc -->
- [Variable Iteration](#variable-iteration)
- [Selecting Cases](#selecting-cases)
- [Subroutines](#subroutines)
  - [Defining A Subroutine](#defining-a-subroutine)
  - [Calling A Subroutine](#calling-a-subroutine)
  - [Parameters](#parameters)
- [Functions](#functions)
  - [Defining A Function](#defining-a-function)
  - [Calling A Function](#calling-a-function)
  - [Function Parameters](#function-parameters)
  - [Return Values](#return-values)

## Guide <!-- omit from toc -->

### Variable Iteration

To increase the value of a variable `x` by 1, you may have been taught the following: 

```
x = x + 1
```

But, instead of this, you can do:
```
x += 1
```

This works for other operations too:
```
x -= 1 'subtracts 1 from x
x *= 1 'multiplies x by 1
x /= 1 'divides x by 1
```
Of course, some of these are completely useless (such as dividing by 1), but you can replace 1 with any other value, including another variable:
```
x += y
```

### Selecting Cases
`If` statements are incredibly powerful. However, they can be time consuming to type out, and awkward to use, with long statements. So, in some instances, a `Select Case` block is more suited for the job. Take this `If` statement:
```
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
```
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

However, if statements can be more powerful

### Subroutines
Subroutines help massively with organising your code. I have seen code before with a `Sub Main` that is over 300 lines long and it *physically hurts* when people tell me they don't know how subroutines work. They are incredibly simple and people seem to overcomplicate them a lot.


#### Defining A Subroutine
To define a subroutine, simply go outside of `Sub Main` but inside `Module Program`, and type in the following code:

```
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
```
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
```
    Sub Main()
        Dim x As String = 10
        MyNewSub()
    End Sub
```
and I wanted to use my new subroutine to, say, multiply that value by 2. But there's a problem: variables created within subroutines are stuck inside the sub they are in. Think of it as if the `Sub` and `End Sub` blocks act as a 'wall' around the variable; it can't escape. Unless, of course, you give it a hole that is the perfect shape & size for it to fit through. This is effectively what parameters are.
So, to pass over my variable `x` to my new subroutine, I would do this:

```
    Sub Main()
        MyNewSub(x)
    End Sub

    Sub MyNewSub()
        'Do Something
    End Sub
```

But there's another problem here: **my new sub is not expecting a value**. So, this is where the second part of parameters come in. What I'd do in this scenario is the following:

```
    Sub Main()
        MyNewSub(x)
    End Sub

    Sub MyNewSub(x As Integer)
        x *= 2 'See the "variable iteration" chapter of this guide for an explaination of this line
    End Sub
```

Notice how I specify the data type in the part where `MyNewSub` expects a variable. This is important because if you don't do this, VB doesn't know if it will have a string, integer, boolean, date, time, or something else, and it doesn't like that.
Also notice how when I am calling the subroutine, I don't need to specify the data type. This, again, is important.

I can then use this data **inside the subroutine** to output or calculate more things. However, I cannot normally pass the multiplied value back to the original subroutine using regular subroutines. To do that, I would have to use a [function](#functions)

### Functions
Functions are just [subroutines](#subroutines) where a value is returned to the place where the function was called from. In most programming languages, subroutines are just functions that don't return a value, but for some reason the VB developers felt the need to separate the two. So, we have to work with that.


#### Defining A Function
To define a function, it is the same as [defining a subroutine](#defining-a-subroutine), but you write 'Function' rather than 'Sub'. Easy, right? Just like this:

```
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
```
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

```
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

```
    Function MyNewFunction() As Integer 'Notice that this line specifies the data type now
        'Do Something
        Return something
    End Function
```

So yeah, that's functions. They're really not too complicated. You should use them more. They are very helpful and keep your code more organised. Plus, you need to know them for your GCSEs. So it's a win-win situation really.