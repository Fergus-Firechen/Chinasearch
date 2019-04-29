# Python基础(廖)

## 安装Python  
Python的解释器很多，但使用最广泛的还是CPython。如果要和Java或.Net平台交互，最好的办法不是用Jython或IronPython，而是通过**网络调用来交互，确保各程序之间的独立性**。

## Python基础  
python对大小写敏感。

### 数据类型和变量  
整数，浮点数，字符串，布尔值，空值，列表，字典等。
变量：可以是任意数据类型；——动态语言
常量：变量名全部大写。
注：整数无大小限制；浮点数无大小限制，但超出一定范围计inf。

### 字符串和编码  
ASCII编码：127个字符，大小写英文字母、数字、一些符号。
GB2312：中文至少两个字节，增加中文。
Unicode把所有语言都统一到一套编码里。又发展出可变长的编码的UTF-8。

- Python的字符串  
  单字符：ord(str->整数);chr(整数->str)
  str->bytes: str.encode('utf-8')
  bytes->str: str.decode('utf-8', errors='ignore')
- 格式化  

```python
>>> 'Hello, %s' %s 'world'
'Hello, world'
>>> 'Hi, %s, you have $%d.' % ('Mi', 1000)
'Hi, Mi, you have $1000'
```

### 使用list & tuple

list和tuple是Python内置的有序集合，一个可变，一个不可变。

### 条件判断

计算机之所以能做很多自动化的任务，因为它可以自己做条件判断。

`if`语句执行有个特点，它是从上往下判断，如果在某个判断上是`True`，把该判断对应的语句执行后，就忽略掉剩下的`elif`和`else`。

`if`判断条件还可以简写，比如写：

```python
if x:
    print('True')
```

只要`x`是非零数值、非空字符串、非空list等，就判断为`True`，否则为`False`。

 - 再议input

   返回数据类型为str

### 循环  

让计算机做重复任务的有效的方法。

Python的循环有两种：

第一种是for...in循环，依次把list或tuple中的每个元素迭代出来。

第二种循环是while循环，只要条件满足，就不断循环，条件不满足时退出循环。

- break  

  提前退出循环（通常必须配合if语句）

- continue

  跳过当前的这次循环，直接开始下一次循环（通常必须配合if语句）

- *特别注意*，不要滥用`break`和`continue`语句。`break`和`continue`会造成代码执行逻辑分叉过多，容易出错。大多数循环并不需要用到`break`和`continue`语句。

### 使用dict & set

- dict

  在其他语言中也称为map，使用键-值（key-value）存储，具有极快的查找速度。

  和list比较，dict有以下几个特点：

  1. 查找和插入的速度极快，不会随着key的增加而变慢；  
  2. 需要占用大量的内存，内存浪费多。

  而list相反：

  1. 查找和插入的时间随着元素的增加而增加；
  2. 占用空间小，浪费内存很少。

  所以，**dict是用空间来换取时间**的一种方法。dict的key必须是**不可变对象**。

  要避免key不存在的错误，有两种办法，一是通过`in`判断key是否存在：

  ```python
  >>> 'Thomas' in d
  False
  ```

  二是通过dict提供的`get()`方法，如果key不存在，可以返回`None`，或者自己指定的value：

  ```python
  >>> d.get('Thomas')
  >>> d.get('Thomas', -1)
  -1
  ```

- set

  set和dict类似，也是一组key的集合，但不存储value。set可以看成数学意义上的无序和无重复元素的集合。set中必须是**不可变对象**。

  对于不变对象来说，调用对象自身的任意方法，也不会改变该对象自身的内容。相反，这些方法会创建新的对象并返回，这样，就保证了不可变对象本身永远是不可变的。

## 函数

借助抽象，我们不用关心底层的具体计算过程，而直接在更高的层次上思考问题。函数就是最基本的一种代码抽象的方式。

### 调用函数

Python内置了许多函数，可以直接调用。需要注意传入正确的参数。

https://docs.python.org/3/library/functions.html#abs

### 定义函数

定义一个函数要使用`def`语句，依次写出函数名、括号、括号中的参数和冒号`:`，然后，在缩进块中编写函数体，函数的返回值用return语句返回。如无return，默认返回*None*.

**注意**：函数体内部的语句在执行时，一旦执行到`return`时，函数就执行完毕，并将结果返回。因此，函数内部通过条件判断和循环可以实现非常复杂的逻辑。

- 空函数

  ```python
  def nop():
      pass
  ```

- 参数检查

  调用函数时，如果参数不对，Python解释器会自动检查出来，并抛出*TypeError*。

  对参数类型做检查，只允许整数和浮点数类型的参数。数据类型检查可以用内置函数*isinstance()*实现：

  ```python
  def my_abs(x):
      if not isinstance(x, (int, float)):
          raise TypeError('bad operand type')
      if x >= 0:
          return x
      else:
          return -x
  ```

- 返回多个值

  在语法上，返回一个tuple可以省略括号，而多个变量可以同时接收一个tuple，按位置赋给对应的值，所以，Python的函数返回多值其实就是**返回一个tuple**。

### 函数的参数

定义函数的时候，我们把参数的名字和位置确定下来，函数的接口定义就完成了。对于函数的调用者来说，只需要知道如何传递正确的参数，以及函数将返回什么样的值就够了，函数内部的复杂逻辑被封装起来，调用者无需了解。

- 位置参数

  按照位置顺序依次赋给参数。

- 默认参数

  由于我们经常计算x2，所以，完全可以把第二个参数n的默认值设定为2：

  ```python
  def power(x, n=2):
      s = 1
      while n > 0:
          n = n - 1
          s = s * x
      return s
  ```

  这样，当我们调用`power(5)`时，相当于调用`power(5, 2)`。

  设置默认参数时，有几点要注意：

  一是必选参数在前，默认参数在后，否则Python的解释器会报错；

  二是如何设置默认参数：

  当函数有多个参数时，把变化大的参数放前面，变化小的参数放后面。变化小的参数就可以作为默认参数。

  使用默认参数有什么好处？最大的好处是能降低调用函数的难度。

  注：**默认参数必须是不可变对象**

  为什么要设计`str`、`None`这样的不变对象呢？因为不变对象一旦创建，对象内部的数据就不能修改，这样就减少了由于修改数据导致的错误。此外，由于对象不变，多任务环境下同时读取对象不需要加锁，同时读一点问题都没有。我们在编写程序时，如果可以设计一个不变对象，那就尽量设计成不变对象。

- 可变参数

  可变参数允许传入0个或任意个参数， 这些可变参数在函数调用时自动组装成一个tuple。定义函数的可变参数：

  ```python
  def calc(*numbers):
      sum = 0
      for n in numbers:
          sum = sum + n * n
      return sum
  ```

  如果已经有了一个list或tuple，Python允许你在list或tuple前面加一个`*`号，把list或tuple的元素变成可变参数传进去：

  ```python
  >>> nums = [1, 2, 3]
  >>> calc(*nums)
  14
  ```

- 关键字参数

  关键字参数允许你传入0个或任意个含参数名的参数，这些关键字参数在函数内部自动组装为一个dict。

  ```python
  def person(name, age, **kw):
      print('name:', name, 'age:', age, 'other:', kw)
  ```

  函数`person`除了必选参数`name`和`age`外，还接受关键字参数`kw`。在调用该函数时，可以只传入必选参数：

  ```python
  >>> person('Michael', 30)
  name: Michael age: 30 other: {}
  ```

  也可以传入任意个数的关键字参数：

  ```python
  >>> person('Bob', 35, city='Beijing')
  name: Bob age: 35 other: {'city': 'Beijing'}
  >>> person('Adam', 45, gender='M', job='Engineer')
  name: Adam age: 45 other: {'gender': 'M', 'job': 'Engineer'}
  ```

  关键字参数有什么用？它可以扩展函数的功能。试想你正在做一个用户注册的功能，除了用户名和年龄是必填项外，其他都是可选项，利用关键字参数来定义这个函数就能满足注册的需求。

  如果已经有了一个dict，可以直接写成：

  ```python
  >>> extra = {'city': 'Beijing', 'job': 'Engineer'}
  >>> person('Jack', 24, **extra)
  name: Jack age: 24 other: {'city': 'Beijing', 'job': 'Engineer'}
  ```

- 命名关键字参数

  **命名关键字参数必须传入参数名**，否则将被视为位置参数，报错*TypeError*。

  如果要限制关键字参数的名字，就可以用命名关键字参数，例如，只接收`city`和`job`作为关键字参数。这种方式定义的函数如下：

  ```python
  def person(name, age, *, city, job):
      print(name, age, city, job)
  ```

  - 如果函数定义中已经有了一个可变参数，后面跟着的命名关键字参数就不再需要一个特殊分隔符`*`了：

  ```python
  def person(name, age, *args, city, job):
      print(name, age, args, city, job)
  ```

  - 命名关键字参数可以有缺省值，从而简化调用：

  ```python
  def person(name, age, *, city='Beijing', job):
      print(name, age, city, job)
  ```

  由于命名关键字参数`city`具有默认值，调用时，可不传入`city`参数：

  ```python
  >>> person('Jack', 24, job='Engineer')
  Jack 24 Beijing Engineer
  ```

  - 使用命名关键字参数时，要特别注意，如果没有可变参数，就必须加一个`*`作为特殊分隔符。如果缺少`*`，Python解释器将无法识别位置参数和命名关键字参数：

  ```python
  def person(name, age, city, job):
      # 缺少 *，city和job被视为位置参数
      pass
  ```

- 参数组合

  5种参数都可以组合使用时需注意，参数定义的顺序必须是：**必选参数、默认参数、可变参数、命名关键字参数和关键字参数**。

  比如定义一个函数，包含上述若干种参数：

  ```python
  def f1(a, b, c=0, *args, **kw):
      print('a =', a, 'b =', b, 'c =', c, 'args =', args, 'kw =', kw)
  
  def f2(a, b, c=0, *, d, **kw):
      print('a =', a, 'b =', b, 'c =', c, 'd =', d, 'kw =', kw)
  ```

  在函数调用的时候，Python解释器自动按照参数位置和参数名把对应的参数传进去。

  ```python
  >>> f1(1, 2)
  a = 1 b = 2 c = 0 args = () kw = {}
  >>> f1(1, 2, c=3)
  a = 1 b = 2 c = 3 args = () kw = {}
  >>> f1(1, 2, 3, 'a', 'b')
  a = 1 b = 2 c = 3 args = ('a', 'b') kw = {}
  >>> f1(1, 2, 3, 'a', 'b', x=99)
  a = 1 b = 2 c = 3 args = ('a', 'b') kw = {'x': 99}
  >>> f2(1, 2, d=99, ext=None)
  a = 1 b = 2 c = 0 d = 99 kw = {'ext': None}
  ```

  最神奇的是通过一个tuple和dict，你也可以调用上述函数：

  ```python
  >>> args = (1, 2, 3, 4)
  >>> kw = {'d': 99, 'x': '#'}
  >>> f1(*args, **kw)
  a = 1 b = 2 c = 3 args = (4,) kw = {'d': 99, 'x': '#'}
  >>> args = (1, 2, 3)
  >>> kw = {'d': 88, 'x': '#'}
  >>> f2(*args, **kw)
  a = 1 b = 2 c = 3 d = 88 kw = {'x': '#'}
  ```

  所以，对于任意函数，都可以通过类似`func(*args, **kw)`的形式调用它，无论它的参数是如何定义的。

  **虽然可以组合多达5种参数，但不要同时使用太多的组合，否则函数接口的可理解性很差。**

### 递归函数

**使用递归函数的优点是逻辑简单清晰，缺点是过深的调用会导致栈溢出。**

在函数内部，可以调用其他函数。如果一个函数在内部调用自身本身，这个函数就是递归函数。

计算n!=1x2x3x4...xn， `fact(n)`用递归的方式写出来就是：

```python
def fact(n):
    if n==1:
        return 1
    return n * fact(n - 1)
```

递归函数的优点是**定义简单，逻辑清晰**。

**使用递归函数需要注意防止栈溢出。**在计算机中，函数调用是通过栈（stack）这种数据结构实现的，每当进入一个函数调用，栈就会加一层栈帧，每当函数返回，栈就会减一层栈帧。由于栈的大小不是无限的，所以，递归调用的次数过多，会导致栈溢出。

解决递归调用栈溢出的方法是通过**尾递归**优化，事实上尾递归和循环的效果是一样的，所以，把循环看成是一种特殊的尾递归函数也是可以的。

**尾递归是指，在函数返回的时候，调用自身本身，并且，return语句不能包含表达式。这样，编译器或者解释器就可以把尾递归做优化，使递归本身无论调用多少次，都只占用一个栈帧，不会出现栈溢出的情况。**

上面的`fact(n)`函数由于`return n * fact(n - 1)`引入了乘法表达式，所以就不是尾递归了。要改成尾递归方式，需要多一点代码，主要是要把每一步的乘积传入到递归函数中：

```python
def fact(n):
    return fact_iter(n, 1)

def fact_iter(num, product):
    if num == 1:
        return product
    return fact_iter(num - 1, num * product)
```

可以看到，`return fact_iter(num - 1, num * product)`仅返回递归函数本身，`num - 1`和`num * product`在函数调用前就会被计算，不影响函数调用。

遗憾的是，**大多数编程语言没有针对尾递归做优化，Python解释器也没有做优化**，所以，即使把上面的`fact(n)`函数改成尾递归方式，也会导致栈溢出。

## 高级特性

在Python中，代码不是越多越好，而是越少越好。代码不是越复杂越好，而是越简单越好。1行代码能实现的功能，决不写5行代码。

### Slice

取一个list、tuple、str的部分元素是非常常见的操作。

### Iteration

给定一个list或tuple，可以通过`for`循环来遍历这个list或tuple，这种遍历我们称为**迭代**（Iteration）。

只要是**可迭代对象**，无论有无下标，都可以迭代。

- 如何判断一个对象是可迭代对象呢？方法是通过collections模块的Iterable类型判断：

```python
>>> from collections import Iterable
>>> isinstance('abc', Iterable) # str是否可迭代
True
>>> isinstance([1,2,3], Iterable) # list是否可迭代
True
>>> isinstance(123, Iterable) # 整数是否可迭代
False
```

- Python内置的`enumerate`函数可以把一个list变成索引-元素对，这样就可以在`for`循环中同时迭代索引和元素本身：

```python
>>> for i, value in enumerate(['A', 'B', 'C']):
...     print(i, value)
...
0 A
1 B
2 C
```

### List Comprehensions

Python内置的非常简单却强大的可以用来创建list的生成式。

如果要生成`[1x1, 2x2, 3x3, ..., 10x10]`怎么做？列表生成式则可以用一行语句代替循环生成上面的list：

```python
>>> [x * x for x in range(1, 11)]
[1, 4, 9, 16, 25, 36, 49, 64, 81, 100]
```

还可以使用两层循环，可以生成全排列：

```python
>>> [m + n for m in 'ABC' for n in 'XYZ']
['AX', 'AY', 'AZ', 'BX', 'BY', 'BZ', 'CX', 'CY', 'CZ']
```

三层和三层以上的循环就很少用到了。

### generator

在Python中，为节约内存，将列表元素推算的方法保存下来，一边循环一边计算的机制，称为**生成器：generator**。

- 第一种方法很简单，只要把一个列表生成式的`[]`改成`()`，就创建了一个generator：

```python
>>> L = [x * x for x in range(10)]
>>> L
[0, 1, 4, 9, 16, 25, 36, 49, 64, 81]
>>> g = (x * x for x in range(10))
>>> g
<generator object <genexpr> at 0x1022ef630>
```

如果要一个一个打印出来，可以通过`next()`函数获得generator的下一个返回值：

```python
>>> next(g)
0
>>> next(g)
1
.
.
>>> next(g)
81
>>> next(g)  # 没有更多了
Traceback (most recent call last):
  File "<stdin>", line 1, in <module>
StopIteration
```

一般不调用`next(g)`，正确的方法是使用`for`循环，因为generator也是可迭代对象。

- 用函数来实现

斐波拉契数列，用函数把它打印出来却很容易：

```python
def fib(max):
    n, a, b = 0, 0, 1
    while n < max:
        print(b)
        a, b = b, a + b
        n = n + 1
    return 'done'

>>> fib(6)
1
1
2
3
5
8
'done'
```

把`fib`函数变成generator，只需要把`print(b)`改为`yield b`即可。如果一个函数定义中包含`yield`关键字，那么这个函数就不再是一个普通函数，而是一个generator：

```python
def fib(max):
    n, a, b = 0, 0, 1
    while n < max:
        yield b
        a, b = b, a + b
        n = n + 1
    return 'done'
    
>>> f = fib(6)
>>> f
<generator object fib at 0x104feaaa0>
```

最难理解的就是**generator和函数的执行流程不一样**。函数是顺序执行，遇到`return`语句或者最后一行函数语句就返回。而变成generator的函数，在每次调用`next()`的时候执行，遇到`yield`语句返回，再次执行时从上次返回的`yield`语句处继续执行。

同样的，把函数改成generator后，我们基本上从来不会用`next()`来获取下一个返回值，而是直接使用`for`循环来迭代。

但是用`for`循环调用generator时，发现拿不到generator的`return`语句的返回值。如果想要拿到返回值，必须捕获`StopIteration`错误，返回值包含在`StopIteration`的`value`中：

```python
>>> g = fib(6)
>>> while True:
...     try:
...         x = next(g)
...         print('g:', x)
...     except StopIteration as e:
...         print('Generator return value:', e.value)
...         break
...
g: 1
g: 1
g: 2
g: 3
g: 5
g: 8
Generator return value: done
```

### Iterator

- 可迭代对象

可以直接作用于`for`循环的数据类型有以下几种：

一类是集合数据类型，如`list`、`tuple`、`dict`、`set`、`str`等；

一类是`generator`，包括生成器和带`yield`的generator function。

这些可以直接作用于`for`循环的对象统称为可迭代对象：`Iterable`。

- 迭代器

可以被`next()`函数调用并不断返回下一个值的对象称为迭代器：`Iterator`。它们表示一个惰性计算的序列。

生成器都是`Iterator`对象。把`list`、`dict`、`str`等`Iterable`变成`Iterator`可以使用`iter()`函数：

```python
>>> isinstance(iter([]), Iterator)
True
>>> isinstance(iter('abc'), Iterator)
True
```

`Iterator`甚至可以表示一个无限大的数据流，例如全体自然数。

## Functional Programming

函数是Python内建支持的一种封装，我们通过把大段代码拆成函数，通过一层一层的函数调用，就可以把复杂任务分解成简单的任务，这种分解可以称之为**面向过程的程序设计**。函数就是面向过程的程序设计的**基本单元**。

函数式编程的一个特点就是，允许把函数本身作为参数传入另一个函数，还允许返回一个函数！

Python对函数式编程提供部分支持。由于Python允许使用变量，因此，Python不是纯函数式编程语言。

### Higher-order function

- 变量可以指向函数

  ```python
  abs
  Out[13]: <function abs(x, /)>
  
  abs(-1)
  Out[14]: 1
  
  a = abs
  
  a
  Out[16]: <function abs(x, /)>
  
  a(-1)
  Out[17]: 1
  ```

  结论：函数本身也可以赋值给变量，即：变量可以指向函数。

- 函数名也是变量

- 传入函数

既然变量可以指向函数，函数的参数能接收变量，那么一个函数就可以接收另一个函数作为参数，这种函数就称之为**高阶函数**。

#### map/reduce

- map

`map()`函数接收两个参数，一个是函数，一个是`Iterable`，`map`将传入的函数依次作用到序列的每个元素，并把结果作为新的`Iterator`返回。

```python
>>> def f(x):
...     return x * x
...
>>> r = map(f, [1, 2, 3, 4, 5, 6, 7, 8, 9])
>>> list(r)
[1, 4, 9, 16, 25, 36, 49, 64, 81]
```

- reduce

`reduce`把一个函数作用在一个序列`[x1, x2, x3, ...]`上，这个函数必须接收两个参数，`reduce`把结果继续和序列的下一个元素做累积计算。

```python
from functools import reduce

DIGITS = {'0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9}

def char2num(s):
    return DIGITS[s]

def str2int(s):
    return reduce(lambda x, y: x * 10 + y, map(char2num, s))

>>> str2int('135')
Out[19]: 135
```

#### filter

`filter()`函数用于过滤序列。`filter()`接收一个函数和一个序列，把传入的函数依次作用于每个元素，然后根据返回值是`True`还是`False`决定保留还是丢弃该元素。

```python
def is_odd(n):
    return n % 2 == 1

list(filter(is_odd, [1, 2, 4, 5, 6, 9, 10, 15]))
# 结果: [1, 5, 9, 15]
```

#### sorted

```python
>>> sorted(['bob', 'about', 'Zoo', 'Credit'], key=str.lower, reverse=True)
['Zoo', 'Credit', 'bob', 'about']
```

默认情况下，对字符串排序，是按照ASCII的大小比较的，由于`'Z' < 'a'`，结果，大写字母`Z`会排在小写字母`a`的前面。

用`sorted()`排序的关键在于实现一个映射函数。

### 返回函数(&Closure)

- 函数作为返回值

```python
def lazy_sum(*args):
    def sum():
        ax = 0
        for n in args:
            ax = ax + n
        return ax
    return sum
>>> f = lazy_sum(1, 3, 5, 7, 9)
>>> f
<function lazy_sum.<locals>.sum at 0x101c6ed90>
>>> f()
25
# 每次调用都会返回一个新的函数，即使传入相同的参数结果都互不影响
>>> f1 = lazy_sum(1, 3, 5, 7, 9)
>>> f2 = lazy_sum(1, 3, 5, 7, 9)
>>> f1==f2
False
```

在函数`lazy_sum`中又定义了函数`sum`，并且，内部函数`sum`可以引用外部函数`lazy_sum`的参数和局部变量，当`lazy_sum`返回函数`sum`时，相关参数和变量都保存在返回的函数中，这种称为“闭包（Closure）”的程序结构拥有极大的威力。

- Closure

  - 返回的函数在其定义内部引用了局部变量`args`，所以，当一个函数返回了一个函数后，其内部的局部变量还被新函数引用，所以，闭包用起来简单，实现起来可不容易。
  - 返回的函数并没有立刻执行，而是直到调用了`f()`才执行。
  - 返回闭包时牢记一点：**返回函数不要引用任何循环变量，或者后续会发生变化的变量**。

  如果一定要引用循环变量怎么办？方法是再创建一个函数，用该函数的参数绑定循环变量当前的值，无论该循环变量后续如何更改，已绑定到函数参数的值不变：

  ```python
  def count():
      def f(j):
          def g():
              return j*j
          return g
      fs = []
      for i in range(1, 4):
          fs.append(f(i)) # f(i)立刻被执行，因此i的当前值被传入f()
      return fs
  >>> f1, f2, f3 = count()
  >>> f1()
  1
  >>> f2()
  4
  >>> f3()
  9
  ```

  缺点是代码较长，可利用lambda函数缩短代码。

### 匿名函数

Python对匿名函数的支持有限，只有一些简单的情况下可以使用匿名函数。

当传入函数时，有时不需要显示的定义函数，直接传入匿名函数。

匿名函数有个限制，就是**只能有一个表达式**，不用写`return`，返回值就是该表达式的结果。

1. 可以把匿名函数赋值给一个变量，再利用变量来调用该函数

```python
>>> f = lambda x: x * x
>>> f
<function <lambda> at 0x101c6ef28>
>>> f(5)
25
```

2. 也可以把匿名函数作为返回值返回

```python
def build(x, y):
    return lambda: x * x + y * y
```

### 装饰器

