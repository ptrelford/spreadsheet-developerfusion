#if SILVERLIGHT
namespace Silverlight.Spreadsheet

open Spreadsheet

type App() as app =
    inherit System.Windows.Application()
    do app.Startup.Add(fun _ -> app.RootVisual <- createGrid sheet)
#else
module App

open System.Windows
open Spreadsheet

[<System.STAThread>]
do  let win = new Window(Title="Spreadsheet", Content=createGrid sheet)
    (new Application()).Run(win) |> ignore
#endif