# stock-analysis
Module 2, VBA, Class repository

Mock-project in which we help our friend's parent analyze green companies to diversify their stock portfolio

For your written analysis, be sure to use complete and coherent sentences. Your written analysis should contain three sections, which cover the following:

----------------------------------------

1) Overview of Project: Explain the purpose of this analysis.

This analysis was to help a mock-parent duo of our mock-friend analyze both the stock they have their life savings invested in, and a few 'green' alternatives
 
 ---------------------------------   

2) Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
    
- Unfortunately I did not save screenshots of the runtime of the code before refactoring, but it was roughly 26 seconds for either sheet.
        Here are the performances of the refactored code:        
![VBA_Challenge_2017](https://user-images.githubusercontent.com/21095468/123361426-a6e56800-d534-11eb-8d8b-d61a4b5505f5.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/21095468/123361430-a9e05880-d534-11eb-8d8b-9ee5d99cb308.png)

  --------------------------------------  
    
3) Summary: In a summary statement, address the following questions.
       
       
3a. What are the advantages or disadvantages of refactoring code?
        
- The advantages are the massive decrease in runtime. Especially for scaling the code up to larger samples, efficient processing can save huge amounts of time.         The major disadvantage I percieve is that, while some reactoring may simplify code, in this case, and likely others, we used more advanced techniques to increase   efficiency, which has the drawback of making the code less readable for others, or at worst incomprehensible to someone not versed in the technique being used.
       
       
      
3b. How do these pros and cons apply to refactoring the original VBA script?
   
- The script runs about 30-40x faster, which is a very significant improvement. Unfortunately the arrays may be too complicated for a user such as our friend, or especially his parents, to understand.
