# utl-excel-fixing-bad-formatting-using-passthru
Excel fixing bad formatting using passthru.
    Excel fixing bad formatting using passthru                                                           
                                                                                                         
    github                                                                                               
    https://github.com/rogerjdeangelis/utl-excel-fixing-bad-formatting-using-passthru                    
                                                                                                         
    SAS  Forum                                                                                           
    https://communities.sas.com/t5/SAS-Programming/PROC-IMPORT-XLS-engine-mixed-columns/m-p/502853       
                                                                                                         
    Other excel repositories                                                                             
    https://tinyurl.com/ybnm6azh                                                                         
    https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language= 
                                                                                                         
    INPUT                                                                                                
    =====                                                                                                
                                                                                                         
      You have to use the excel sheet in github                                                          
                                                                                                         
       https://tinyurl.com/ya3alq9c                                                                      
       https://github.com/rogerjdeangelis/utl-excel-fixing-bad-formatting-using-passthru/blob/master/    
       utl_excel_fixing_bad_formatting_using_passthru.xls                                                
                                                                                                         
          +--------------------------------------+                                                       
          |     A      |    B       |     C      |                                                       
          +--------------------------------------+                                                       
       1  | FOO1       |   FOO2     |   FOO3     |                                                       
          +------------+------------+------------+                                                       
       2  |  1234567890| 12345E+13  |1234567890  |                                                       
          +------------+------------+------------+                                                       
       3  |    9123456 | 23456E+12  |8974658     |                                                       
          +------------+------------+------------+                                                       
                                                                                                         
    EXAMPLE OUTPUT (SAS Dataset)                                                                         
                                                                                                         
     WORK.WANT total obs=5                                                                               
                                                                                                         
     Obs        FOO1              FOO2              FOO3                                                 
                                                                                                         
      1    12345678901234    12345678901234    12345678901234                                            
      2    1234567890123     1234567890123     1234567890123                                             
      3    123456789012      123456789012      123456789012                                              
      4    12345678901       12345678901       12345678901                                               
      5    1234567890        1234567890        1234567890                                                
                                                                                                         
                                                                                                         
    PROCESS                                                                                              
    =======                                                                                              
                                                                                                         
    proc sql dquote=ansi;                                                                                
     connect to excel                                                                                    
        (Path="d:/xls/bignumber.xls" );                                                                  
        create                                                                                           
            table want as                                                                                
        select                                                                                           
            *                                                                                            
            from connection to Excel                                                                     
            (                                                                                            
             Select                                                                                      
                format(foo1,'################') as foo1                                                  
               ,format(foo2,'################') as foo2                                                  
               ,format(foo3,'################') as foo3                                                  
             from                                                                                        
               [foo]                                                                                     
            );                                                                                           
        disconnect from Excel;                                                                           
    Quit;                                                                                                
                                                                                                         
    OUTPUT                                                                                               
    ======                                                                                               
                                                                                                         
    see above                                                                                            
                                                                                                         
    *                _              _       _                                                            
     _ __ ___   __ _| | _____    __| | __ _| |_ __ _                                                     
    | '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |                                                    
    | | | | | | (_| |   <  __/ | (_| | (_| | || (_| |                                                    
    |_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|                                                    
                                                                                                         
    ;                                                                                                    
                                                                                                         
    use                                                                                                  
                                                                                                         
    https://tinyurl.com/ya3alq9c                                                                         
                                                                                                         
