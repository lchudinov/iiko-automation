$a = "hello"

function hello() {
    $b = if ($a -eq "hello") { "it's hello" } else {"it's not hello"} 
    Write-Host $b
}

hello