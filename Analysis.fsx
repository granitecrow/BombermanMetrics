#r "c:\\users\\jeffrey.cusato.jade\\documents\\visual studio 2015\\Projects\\BombermanAnalysis\\packages\\ExcelProvider.0.8.0\\lib\\ExcelProvider.dll"
open FSharp.ExcelProvider
open System

// need to figure out the whole time thing...
type TimeRemaining(dt:DateTime) = 
  member x.Second = dt.Second
  member x.Minute = dt.Minute

type DataTypesTest = ExcelFile<"C:\\Users\\jeffrey.cusato.JADE\\Documents\\Visual Studio 2015\\Projects\\BombermanAnalysis\\BombermanAnalysis\\bomberman_metrics.xlsx", SheetName="PlayerStats">
let file = new DataTypesTest()

let headerCols = file.Data |> Seq.head
let rows = file.Data |> Seq.toArray

//example of yielding a custom sequence
//let ListOfWinnersAndTheirStartLocations =
//    [for r in rows do
//        if r.Winner = "yes" then
//            yield r.Player, r.StartingRing, r.StartingQuadrant]

//
// HELPER FUNCTIONS
//
let GroupAndCount elems =
    elems
    |> Seq.groupBy (fun x -> x)
    |> Seq.map (fun pair ->
                    let killtype = fst pair
                    let count = snd pair |> Seq.length
                    killtype, count)

let PrintSeq data = Seq.iter (printfn "%A") data
//
// END HELPER FUNCTIONS
//

// HOW PLAYER DIED
let KilledUsing player week =
    rows
    |> Seq.filter (fun p -> p.Player = player && p.Week = week)
    |> Seq.map (fun p -> p.KilledUsing)
    |> Seq.filter (fun elem -> not (String.IsNullOrEmpty elem))
    |> GroupAndCount
    |> PrintSeq

let WhoKilledMe player week =
    rows
    |> Seq.filter (fun p -> p.Player = player && p.Week = week)
    |> Seq.map (fun p -> p.KilledBy)
    |> Seq.filter (fun elem -> not (String.IsNullOrEmpty elem))
    |> GroupAndCount
    |> PrintSeq

let CursedDeaths player =
    rows
    |> Seq.filter(fun p -> p.Player = player)
    |> Seq.map (fun p -> p.CurseDeath)
    |> Seq.filter(fun elem -> String.IsNullOrEmpty elem)
    |> Seq.length

// HOW PLAYER KILLED
let TotalSuicides player =
    rows
    |> Seq.filter (fun p -> p.Player = player)
    |> Seq.where (fun p -> p.KilledBy = p.Player)
    |> Seq.length

let SuicidesForPlayerInWeek player week =
    rows
    |> Seq.filter (fun p -> p.Player = player && p.Week = week)
    |> Seq.where (fun p -> p.KilledBy = p.Player)
    |> Seq.length

let TotalKills player =
    rows
    |> Seq.filter (fun p -> p.KilledBy = player)
    |> Seq.length

let KillsForPlayerInWeek player week =
    rows
    |> Seq.filter (fun p -> p.Week = week)
    |> Seq.where (fun p -> p.KilledBy = player)
    |> Seq.length

let TotalOpponentKills player = (TotalKills player) - (TotalSuicides player)
let OpponentKillsInWeek player week = (KillsForPlayerInWeek player week) - (SuicidesForPlayerInWeek player week)



// START POSITION
let StartPositionWhenWin player week =
    rows
    |> Seq.filter (fun p -> p.Player = player && p.Week = week && p.Winner = "yes")
    |> Seq.map (fun p -> p.StartingRing, p.StartingQuadrant)
    |> GroupAndCount
    |> PrintSeq

// SIMPLE MATCH-WIN STATS
let MatchesPerWeek player =
    rows
    |> Seq.filter(fun p -> p.Player = player)
    |> Seq.map(fun p -> p.Week)
    |> GroupAndCount
    |> PrintSeq

let TotalMatchesPlayed player =
    rows
    |> Seq.filter(fun p -> p.Player = player)
    |> Seq.map(fun p -> p.Week)
    |> GroupAndCount
    |> Seq.map (fun (_, b) -> b)
    |> Seq.sum

let TotalWins player = 
    rows
    |> Seq.filter(fun p -> p.Player = player && p.Winner = "yes")
    |> Seq.length

let VictoryRate player = (float(TotalWins player)) / (float(TotalMatchesPlayed player))



let AverageRemaingTime player week =
    rows
    |> Seq.filter (fun p -> p.Player = player && p.Week = week)
//    |> Seq.filter (fun elem -> not (Double.IsNaN elem))
    |> Seq.map (fun p -> p.TimeRemainingAtDeath.Minute * 60 + p.TimeRemainingAtDeath.Second)



// when you win, how much do you kill?

// Metrics for player
// You have played 2 weeks and have 1 victory
// You played 30 matches and won 15 yielding a win rate of 0.50
// You have been directly responsible for 30 deaths
// although 20 are suicides!
// Here is the total ways you have died
// 


let Bombermetric player =
    printfn "Metrics for %A..." player
    printfn "You have played %A matches" (TotalMatchesPlayed player)
    printfn "and you have won %A of them" (TotalWins player)
    printfn "yielding a win rate of %A" (VictoryRate player)
    printfn "You have been directly responsible for %A deaths." (TotalKills player)
    printfn "although %A were suicides!" (TotalSuicides player)
    


     