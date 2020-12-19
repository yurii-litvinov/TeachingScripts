open System

if fsi.CommandLineArgs.Length = 2 then
   let param = fsi.CommandLineArgs.[1]
   let repoName = param.Split '/' |> Array.rev |> Seq.take 2 |> Seq.rev |> Seq.map (fun x -> x.Replace(".git", "")) |> fun s -> String.Join("-", s)
   printfn "Cloning into %s" repoName
   Diagnostics.Process.Start("CMD.exe", "/C git clone " + param + " " + repoName).WaitForExit()
   0
else
   printfn "%s" "Expecting repository URL as an argument"
   1