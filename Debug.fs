[<RequireQualifiedAccess>]
module Debug

open System
open System.Reflection


let printPublicProperties (obj: obj) =
    let objType = obj.GetType()

    let properties =
        objType.GetProperties(BindingFlags.Public ||| BindingFlags.Instance)

    for prop in properties do
        let value = prop.GetValue(obj, null)
        printfn "Property: %s, Value: %A" prop.Name value

let pause () =
    Console.WriteLine("回车继续")
    Console.ReadLine() |> ignore