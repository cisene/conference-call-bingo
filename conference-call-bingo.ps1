#
# Source: unknown
#
#
#
#

$Code = @"
using System;
using System.Security.Cryptography;

public class BingoSquare
{
    [Flags]
    public enum BingoType
    {
        None = 0,
        Horizontal = 1,
        Vertical = 2,
        TopLeftDiagonal = 4,
        TopRightDiagonal = 8
    }
    [Flags]
    private enum DiagonalDirection
    {
        None = 0,
        TopLeftToBottomRight = 1,
        TopRightToBottomLeft = 2
    }

    private static readonly RandomNumberGenerator rng = RandomNumberGenerator.Create();

    public static ulong GetPRNGNumber(ulong max)
    {
        var bytes = new byte[8];
        rng.GetBytes(bytes);
        return BitConverter.ToUInt64(bytes, 0) % max;
    }
    public static int GetPRNGNumber(int max)
    {
        if (max < 1)
            throw new ArgumentOutOfRangeException("max");
        return (int)GetPRNGNumber((ulong)max);
    }
    private short _dimensions;

    private int[] _rows;
    private int[] _cols;
    private bool[,] _matrix;
    private int[] _diagonals = new int[2]; // Downslope, Upslope

    public BingoSquare(short dimensions)
    {
        if (dimensions < 0)
            throw new ArgumentOutOfRangeException("dimensions");
        _dimensions = dimensions;
        _matrix = new bool[dimensions, dimensions];
        _rows = new int[dimensions];
        _cols = new int[dimensions];
        _diagonals = new int[2]; // Downslope, Upslope
    }

    private bool IsOnDiagonal(short row, short col, out DiagonalDirection direction)
    {
        var topleft = row == col;
        var topright = row + col == _dimensions - 1;
        direction = DiagonalDirection.None;
        if (topleft)
            direction |= DiagonalDirection.TopLeftToBottomRight;
        if (topright)
            direction |= DiagonalDirection.TopRightToBottomLeft;
        return topleft || topright;
    }

    public BingoType Add(short row, short col)
    {
        if (row > _dimensions - 1 || col > _dimensions - 1)
            throw new ArgumentOutOfRangeException();
        var type = BingoType.None;
        if (!_matrix[row, col])
        {
            _matrix[row, col] = true;
            if (++_rows[row] == _dimensions)
            {
                type |= BingoType.Horizontal;
            }
            if (++_cols[col] == _dimensions)
            {
                type |= BingoType.Vertical;
            }
            DiagonalDirection direction;
            if (IsOnDiagonal(row, col, out direction))
            {
                if ((direction & DiagonalDirection.TopLeftToBottomRight) == DiagonalDirection.TopLeftToBottomRight)
                {
                    // Downslope
                    if (++_diagonals[0] == _dimensions)
                    {
                        type |= BingoType.TopLeftDiagonal;
                    }
                }
                if ((direction & DiagonalDirection.TopRightToBottomLeft) == DiagonalDirection.TopRightToBottomLeft)
                {
                    // Upslope
                    if (++_diagonals[1] == _dimensions)
                    {
                        type |= BingoType.TopRightDiagonal;
                    }
                }
            }
        }
        return type;
    }
    public void Remove(short row, short col)
    {
        if (row > _dimensions - 1 || col > _dimensions - 1)
            throw new ArgumentOutOfRangeException();
        if (_matrix[row, col])
        {
            _matrix[row, col] = false;
            --_rows[row];
            --_cols[col];
            DiagonalDirection direction;
            if (IsOnDiagonal(row, col, out direction))
            {
                if ((direction & DiagonalDirection.TopLeftToBottomRight) == DiagonalDirection.TopLeftToBottomRight)
                {
                    // Downslope
                    --_diagonals[0];
                }
                if ((direction & DiagonalDirection.TopRightToBottomLeft) == DiagonalDirection.TopRightToBottomLeft)
                {
                    // Upslope
                    --_diagonals[1];
                }
            }
        }
    }
    public BingoType IsBingo(short row, short col)
    {
        if (row > _dimensions - 1 || col > _dimensions - 1)
            throw new ArgumentOutOfRangeException();
        var type = BingoType.None;
        if (_rows[row] == _dimensions)
        {
            type |= BingoType.Horizontal;
        }
        if (_cols[col] == _dimensions)
        {
            type |= BingoType.Vertical;
        }
        DiagonalDirection direction;
        if (IsOnDiagonal(row, col, out direction))
        {
            if ((direction & DiagonalDirection.TopLeftToBottomRight) == DiagonalDirection.TopLeftToBottomRight)
            {
                // Downslope
                if (_diagonals[0] == _dimensions)
                {
                    type |= BingoType.TopLeftDiagonal;
                }
            }
            if ((direction & DiagonalDirection.TopRightToBottomLeft) == DiagonalDirection.TopRightToBottomLeft)
            {
                // Upslope
                if (_diagonals[1] == _dimensions)
                {
                    type |= BingoType.TopRightDiagonal;
                }
            }
        }
        return type;
    }
    public bool[,] GetMatrix()
    {
        //var matrix = new bool[_dimensions, _dimensions];
        return _matrix.Clone() as bool[,];
    }
}

"@
Add-Type -TypeDefinition $Code -Language CSharp -ReferencedAssemblies System,mscorlib,System.Security

$Texts = @(
    "Hi, who just joined?",
    "Can you email that to everyone?",
    "x, are you there?",
    "Uh, x, you're still sharing...",
    "Hey guys I have to jump to another call",
    "<Sound of someone typing, possibly with a hammer>",
    "<Loud, painful echo/feedback>",
    "<Child or animal noises>",
    "<Commute PA announcement>",
    "Hi, can you hear me?",
    "No, it's still loading",
    "Next slide, please.",
    "Can everyone go on mute?",
    "I'm sorry, I was on mute.",
    "(For overtalkers) Sorry, go ahead.",
    "Hello? Hello?",
    "So <faded out>, I can <unintelligible> by <cuts out> ok?",
    "Sorry I'm late <insert lame excuse>",
    "I have a hard stop at...",
    "I'm sorry, you cut out there",
    "Can we take this offline?",
    "I'll have to get back to you",
    "<off-topic remark>",
    "Can everyone see my screen?",
    "Sorry, I was having connection issues",
    "I think there's a lag",
    "Sorry, didn't catch that, can you repeat?",
    "Automated announcer voice constantly interrupts speaker"
)

Function Scramble
{
    [Cmdletbinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [String[]]
        $InputObject
    )
    Process
    {
        For ($i = $InputObject.Length - 1; $i -gt 0; $i--)
        {
            $NewIndex = [BingoSquare]::GetPRNGNumber($i)
            $Item = $InputObject[$i]
            $InputObject[$i] = $InputObject[$NewIndex]
            $InputObject[$NewIndex] = $Item
        }
        return $InputObject
    }
}
$ScrambledTexts = Scramble -InputObject $Texts
$Dimensions = 5
$Bingo = [BingoSquare]::new($Dimensions)

Add-Type -AssemblyName System.Windows.Forms
$Form = New-Object -TypeName System.Windows.Forms.Form
$Form.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 770, 800
$Form.Text = "Conference Call Bingo!"

$OffsetX = 25
$OffsetY = 25
$ButtonWidth = 140
$ButtonHeight = 140
$ControlColor = [System.Drawing.Color]::FromKnownColor([System.Drawing.KnownColor]::Control)
$MarkedColor = [System.Drawing.Color]::FromKnownColor([System.Drawing.KnownColor]::YellowGreen)
$TagList = [ordered]@{
    X = 0
    Y = 0
}
$Buttons = New-Object -TypeName "System.Windows.Forms.Button[,]" -ArgumentList 6,6
$Max = $Dimensions*$Dimensions
If ($ScrambledTexts.Length -lt $Max)
{
    $ErrorMsg = "At least {0} texts are required, only {1} are defined, please add more." -f $Max, $ScrambledTexts.Length 
    [System.Windows.Forms.MessageBox]::Show($ErrorMsg, "Too few texts", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}
For ($i = 0; $i -lt $Max; $i++)
{
    $Button = New-Object -TypeName System.Windows.Forms.Button
    $Button.Size = New-Object -TypeName System.Drawing.Size -ArgumentList $ButtonWidth, $ButtonHeight
    $CoordX = $i%$Dimensions
    $CoordY = [math]::Floor($i/$Dimensions)
    $LocationX = $OffsetX + $ButtonWidth*$CoordX
    $LocationY = $OffsetY + $ButtonHeight*$CoordY
    #"X: {0,-5} Y: {1}" -f $CoordX, $CoordY
    $Button.Location = New-Object -TypeName System.Drawing.Point -ArgumentList $LocationX, $LocationY
    $Button.Text = $ScrambledTexts[$i]
    $Button.Name = "{0}|{1}" -f $CoordX, $CoordY
    $Button.BackColor = $ControlColor
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.Tag = New-Object -TypeName System.Drawing.Point -ArgumentList $CoordX, $CoordY
    $Button.Add_Click({
        $b = $this -as [System.Windows.Forms.Button]
        If ($b.BackColor -eq $ControlColor)
        {
            $b.BackColor = $MarkedColor
            $X = $b.Tag.X
            $Y = $b.Tag.Y
            $BingoResult = $Bingo.Add($X, $Y)
            If ($BingoResult -ne [BingoSquare+BingoType]::None)
            {
                $BingoButtons = New-Object -TypeName "System.Collections.Generic.List[System.Windows.Forms.Button]"
                If ($BingoResult.HasFlag([BingoSquare+BingoType]::Horizontal))
                {
                    For ($i = 0; $i -lt $Dimensions; $i++)
                    {
                        $Control = $Buttons[$X, $i]
                        $BingoButtons.Add($Control)
                    }
                }
                If ($BingoResult.HasFlag([BingoSquare+BingoType]::Vertical))
                {
                    For ($i = 0; $i -lt $Dimensions; $i++)
                    {
                        $Control = $Buttons[$i, $Y]
                        $BingoButtons.Add($Control)
                    }
                }
                If ($BingoResult.HasFlag([BingoSquare+BingoType]::TopLeftDiagonal))
                {
                    For ($i = 0; $i -lt $Dimensions; $i++)
                    {
                        $Control = $Buttons[$i, $i]
                        $BingoButtons.Add($Control)
                    }
                }
                If ($BingoResult.HasFlag([BingoSquare+BingoType]::TopRightDiagonal))
                {
                    For ($i = 0; $i -lt $Dimensions; $i++)
                    {
                        $Ynew = $Dimensions - $i - 1
                        $Control = $Buttons[$i, $Ynew]
                        $BingoButtons.Add($Control)
                    }
                }
                Foreach ($BingoButton in $BingoButtons)
                {
                    $BingoButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Red
                    $BingoButton.FlatAppearance.BorderSize = 2
                }
                [System.Windows.Forms.MessageBox]::Show("BINGO!!!")
            }
        }
        Else
        {
            $b.BackColor = $ControlColor
            $X = $b.Tag.X
            $Y = $b.Tag.Y
            $IsBingo = $Bingo.IsBingo($X, $Y)
            If ($IsBingo -ne [BingoSquare+BingoType]::None)
            {
                $BingoButtons = New-Object -TypeName "System.Collections.Generic.List[System.Windows.Forms.Button]"
                If ($IsBingo.HasFlag([BingoSquare+BingoType]::Horizontal))
                {
                    For ($i = 0; $i -lt $Dimensions; $i++)
                    {
                        $Control = $Buttons[$X, $i]
                        $BingoButtons.Add($Control)
                    }
                }
                If ($IsBingo.HasFlag([BingoSquare+BingoType]::Vertical))
                {
                    For ($i = 0; $i -lt $Dimensions; $i++)
                    {
                        $Control = $Buttons[$i, $Y]
                        $BingoButtons.Add($Control)
                    }
                }
                If ($IsBingo.HasFlag([BingoSquare+BingoType]::TopLeftDiagonal))
                {
                    For ($i = 0; $i -lt $Dimensions; $i++)
                    {
                        $Control = $Buttons[$i, $i]
                        $BingoButtons.Add($Control)
                    }
                }
                If ($IsBingo.HasFlag([BingoSquare+BingoType]::TopRightDiagonal))
                {
                    For ($i = 0; $i -lt $Dimensions; $i++)
                    {
                        $Ynew = $Dimensions - $i - 1
                        $Control = $Buttons[$i, $Ynew]
                        $BingoButtons.Add($Control)
                    }
                }
                Foreach ($BingoButton in $BingoButtons)
                {
                    $BingoButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
                    $BingoButton.FlatAppearance.BorderSize = 1
                }
            }
            $BingoResult = $Bingo.Remove($X, $Y)
        }
    })
    $Buttons[$CoordX, $CoordY] = $Button
    $Form.Controls.Add($Button)
}
$Form.ShowDialog()