#if INTERACTIVE
#else
module Spreadsheet
#endif

type token =
    | WhiteSpace
    | Symbol of char
    | OpToken of string
    | RefToken of int * int
    | StrToken of string
    | NumToken of decimal 

let (|Match|_|) pattern input =
    let m = System.Text.RegularExpressions.Regex.Match(input, pattern)
    if m.Success then Some m.Value else None

let toRef (s:string) =
    let col = int s.[0] - int 'A'
    let row = s.Substring 1 |> int
    col, row-1

let toToken = function
    | Match @"^\s+" s -> s, WhiteSpace
    | Match @"^\+|^\-|^\*|^\/"  s -> s, OpToken s
    | Match @"^=|^<>|^<=|^>=|^>|^<"  s -> s, OpToken s   
    | Match @"^\(|^\)|^\,|^\:" s -> s, Symbol s.[0]   
    | Match @"^[A-Z]\d+" s -> s, s |> toRef |> RefToken
    | Match @"^[A-Za-z]+" s -> s, StrToken s
    | Match @"^\d+(\.\d+)?|\.\d+" s -> s, s |> decimal |> NumToken
    | _ -> invalidOp ""

let tokenize s =
    let rec tokenize' index (s:string) =
        if index = s.Length then [] 
        else
            let next = s.Substring index 
            let text, token = toToken next
            token :: tokenize' (index + text.Length) s
    tokenize' 0 s
    |> List.choose (function WhiteSpace -> None | t -> Some t)

type arithmeticOp = Add | Sub | Mul | Div
type logicalOp = Eq | Lt | Gt | Le | Ge | Ne
type formula =
    | Neg of formula
    | ArithmeticOp of formula * arithmeticOp * formula
    | LogicalOp of formula * logicalOp * formula
    | Num of decimal
    | Ref of int * int
    | Range of int * int * int * int
    | Fun of string * formula list

let rec (|Term|_|) = function
    | Sum(f1, (OpToken(LogicOp op))::Sum(f2,t)) -> Some(LogicalOp(f1,op,f2),t)
    | Sum(f1,t) -> Some (f1,t)
    | _ -> None
and (|LogicOp|_|) = function
    | "=" ->  Some Eq | "<>" -> Some Ne
    | "<" ->  Some Lt | ">"  -> Some Gt
    | "<=" -> Some Le | ">=" -> Some Ge
    | _ -> None
and (|Sum|_|) = function
    | Factor(f1, t) ->      
        let rec aux f1 = function        
            | SumOp op::Factor(f2, t) -> aux (ArithmeticOp(f1,op,f2)) t               
            | t -> Some(f1, t)      
        aux f1 t  
    | _ -> None
and (|SumOp|_|) = function 
    | OpToken "+" -> Some Add | OpToken "-" -> Some Sub 
    | _ -> None
and (|Factor|_|) = function  
    | OpToken "-"::Factor(f, t) -> Some(Neg f, t)
    | Atom(f1, ProductOp op::Factor(f2, t)) ->
        Some(ArithmeticOp(f1,op,f2), t)       
    | Atom(f, t) -> Some(f, t)  
    | _ -> None    
and (|ProductOp|_|) = function
    | OpToken "*" -> Some Mul | OpToken "/" -> Some Div
    | _ -> None
and (|Atom|_|) = function      
    | NumToken n::t -> Some(Num n, t)
    | RefToken(x1,y1)::(Symbol ':'::RefToken(x2,y2)::t) -> 
        Some(Range(min x1 x2,min y1 y2,max x1 x2,max y1 y2),t)  
    | RefToken(x,y)::t -> Some(Ref(x,y), t)
    | Symbol '('::Term(f, Symbol ')'::t) -> Some(f, t)
    | StrToken s::Tuple(ps, t) -> Some(Fun(s,ps),t)  
    | _ -> None
and (|Tuple|_|) = function
    | Symbol '('::Params(ps, Symbol ')'::t) -> Some(ps, t)  
    | _ -> None
and (|Params|_|) = function
    | Term(f1, t) ->
        let rec aux fs = function
            | Symbol ','::Term(f2, t) -> aux (fs@[f2]) t
            | t -> fs, t
        Some(aux [f1] t)
    | t -> Some ([],t)

let parse s = 
    tokenize s |> function 
    | Term(f,[]) -> f 
    | _ -> failwith "Failed to parse formula"

let evaluate valueAt formula =
    let rec eval = function
        | Neg f -> - (eval f)
        | ArithmeticOp(f1,op,f2) -> arithmetic op (eval f1) (eval f2)
        | LogicalOp(f1,op,f2) -> if logic op (eval f1) (eval f2) then 0.0M else -1.0M
        | Num d -> d
        | Ref(x,y) -> valueAt(x,y) |> decimal
        | Range _ -> invalidOp "Expected in function"
        | Fun("SUM",ps) -> ps |> evalAll |> List.sum
        | Fun("IF",[condition;f1;f2]) -> 
            if (eval condition)=0.0M then eval f1 else eval f2 
        | Fun(_,_) -> failwith "Unknown function"
    and arithmetic = function
        | Add -> (+) | Sub -> (-) | Mul -> (*) | Div -> (/)
    and logic = function         
        | Eq -> (=)  | Ne -> (<>)
        | Lt -> (<)  | Gt -> (>)
        | Le -> (<=) | Ge -> (>=)
    and evalAll ps =
        ps |> List.collect (function            
            | Range(x1,y1,x2,y2) ->
                [for x=x1 to x2 do for y=y1 to y2 do yield valueAt(x,y) |> decimal]
            | x -> [eval x]            
        )
    eval formula

let references formula =
    let rec traverse = function
        | Ref(x,y) -> [x,y]
        | Range(x1,y1,x2,y2) -> 
            [for x=x1 to x2 do for y=y1 to y2 do yield x,y]
        | Fun(_,ps) -> ps |> List.collect traverse
        | ArithmeticOp(f1,_,f2) | LogicalOp(f1,_,f2) -> 
            traverse f1 @ traverse f2
        | _ -> []
    traverse formula

open System.ComponentModel

type Cell (sheet:Sheet) as cell =
    inherit ObservableObject()
    let mutable value = ""
    let mutable data = ""       
    let mutable formula : formula option = None
    let updated = Event<_>()
    let mutable subscriptions : System.IDisposable list = []   
    let cellAt(x,y) = 
        let (row : Row) = Array.get sheet.Rows y
        let (cell : Cell) = Array.get row.Cells x
        cell
    let valueAt address = (cellAt address).Value
    let eval formula =         
        try (evaluate valueAt formula).ToString()       
        with _ -> "N/A"
    let parseFormula (text:string) =
        if text.StartsWith "="
        then                
            try true, parse (text.Substring 1) |> Some
            with _ -> true, None
        else false, None
    let update newValue generation =
        if newValue <> value then
            value <- newValue
            updated.Trigger generation
            cell.Notify "Value"
    let unsubscribe () =
        subscriptions |> List.iter (fun d -> d.Dispose())
        subscriptions <- []
    let subscribe formula addresses =
        let remember x = subscriptions <- x :: subscriptions
        for address in addresses do
            let cell' : Cell = cellAt address
            cell'.Updated
            |> Observable.subscribe (fun generation ->   
                if generation < sheet.MaxGeneration then
                    let newValue = eval formula
                    update newValue (generation+1)
            ) |> remember
    member cell.Data 
        with get () = data 
        and set (text:string) =
            data <- text        
            cell.Notify "Data"          
            let isFormula, newFormula = parseFormula text               
            formula <- newFormula
            unsubscribe()
            formula |> Option.iter (fun f -> references f |> subscribe f)
            let newValue =
                match isFormula, formula with           
                | _, Some f -> eval f
                | true, _ -> "N/A"
                | _, None -> text
            update newValue 0
    member cell.Value = value
    member cell.Updated = updated.Publish     
and Row (index,colCount,sheet) =
    let cells = Array.init colCount (fun i -> Cell(sheet))
    member row.Cells = cells
    member row.Index = index
and Sheet (colCount,rowCount) as sheet =
    let rows = Array.init rowCount (fun index -> Row(index+1,colCount,sheet))
    member sheet.ColCount = colCount
    member sheet.Rows = rows
    member sheet.MaxGeneration = 1000
and ObservableObject() =
    let propertyChanged = 
        Event<PropertyChangedEventHandler,PropertyChangedEventArgs>()
    member this.Notify name =
        propertyChanged.Trigger(this,PropertyChangedEventArgs name)
    interface INotifyPropertyChanged with
        [<CLIEvent>]
        member this.PropertyChanged = propertyChanged.Publish

#if INTERACTIVE
#r @"C:\Program Files\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\PresentationCore.dll"
#r @"C:\Program Files\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\PresentationFramework.dll"
#r @"C:\Program Files\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Xaml.dll"
#r @"C:\Program Files\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\WindowsBase.dll"
#endif

open System.Windows
open System.Windows.Controls
open System.Windows.Data
open System.Windows.Markup

#if SILVERLIGHT
module XamlReader =
    let Parse (s:string) = XamlReader.Load s
#endif

open System.Windows
open System.Windows.Controls
open System.Windows.Data
open System.Windows.Markup

let createGridColumn i =
    let header = 'A' + char i
    let col = DataGridTemplateColumn(Header=header, Width=DataGridLength(64.0))
    let toDataTemplate s =
        let ns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        sprintf "<DataTemplate xmlns='%s'>%s</DataTemplate>" ns s
        |> XamlReader.Parse :?> DataTemplate
    let path = sprintf "Cells.[%d]" i
    col.CellTemplate <- 
        sprintf "<TextBlock Text='{Binding %s.Value}'/>" path
        |> toDataTemplate
    col.CellEditingTemplate <- 
        sprintf "<TextBox Text='{Binding %s.Data,Mode=TwoWay}'/>" path
        |> toDataTemplate
    col

/// Force row commit after cell edit
let fixWPF (grid:DataGrid) =
    let manualEdit = ref false    
    grid.CellEditEnding.Add(fun e ->
        if not !manualEdit then 
            manualEdit := true
            grid.CommitEdit(DataGridEditingUnit.Row,true) |> ignore
            manualEdit := false
    )

let createGrid (sheet:Sheet) =
    let grid = DataGrid(AutoGenerateColumns=false,HeadersVisibility=DataGridHeadersVisibility.All)
    fixWPF grid
    for i = 0 to sheet.ColCount-1 do createGridColumn i |> grid.Columns.Add
    grid.LoadingRow.Add (fun e ->
        let row = e.Row.DataContext :?> Row
        e.Row.Header <- row.Index
    )
    grid.ItemsSource <- sheet.Rows
    grid

let sheet = Sheet(26,30)

#if INTERACTIVE 
let win = new Window(Title="Spreadsheet", Content=createGrid sheet)
do  win.ShowDialog() |> ignore
#endif
