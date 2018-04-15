$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "ConvertTo-XlAddress" {
    Context "基本機能の確認" {
        It "1 を入力したら A を返すこと" {
            1 | ConvertTo-XlAddress | Should -Be "A"
        }
        It "26 を入力したら Z を返すこと" {
            26 | ConvertTo-XlAddress | Should -Be "Z"
        }
        It "27 を入力したら AA を返すこと" {
            27 | ConvertTo-XlAddress | Should -Be "AA"
        }
        It "数値の配列 1, 2, 3 が渡されたら文字列の配列 A, B, C を返すこと" {
            1..3 | ConvertTo-XlAddress | Should -Be @("A","B","C")
        }
    }
    Context "極端な値が入力されたときのテスト" {
        It "53 を入力したら BA を返すこと" {
            53 | ConvertTo-XlAddress | Should -Be "BA"
        }
        It "703 を入力したら AAA を返すこと"　{
            703 | ConvertTo-XlAddress | Should -Be "AAA"
        }
    }
    Context "オプションの確認" {
        It "数値の配列 1, 2, 3 が渡されたらカンマ区切りの文字列 `"A,B,C`" が出力されること" {
            1..3 | ConvertTo-XlAddress -Delimiter "," | Should -Be "A,B,C"
        }
    }
    Context "異常系のテスト" {
        It "0 以下を入力したらエラーになること" {
            { 0 | ConvertTo-XlAddress } | Should -Throw
        }
        It "数値以外を入力したらエラーになること" {
            { "A" | ConvertTo-XlAddress } | Should -Throw
        }
    }
}
