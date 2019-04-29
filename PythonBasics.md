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

借助抽象，我们才能不关心底层的具体计算过程，而直接在更高的层次上思考问题。函数就是最基本的一种代码抽象的方式。

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

  调用函数时，如果参数个数不对，Python解释器会自动检查出来，并抛出*TypeError*。

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

