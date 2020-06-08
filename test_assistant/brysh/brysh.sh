#!/bin/bash

readonly DEBUG_MODE="ON"

readonly PROGRAM_NAME=$(basename $0)
readonly PROGRAM_PATH=$(cd $(dirname "$0"); pwd)
readonly VERSION="00.00.01"
readonly LAST_UPDATED="2020/06/07"
readonly AUTHOR="tarot ashigaru"

# main
function main()
{
    debug_echo "func:main args:$*"
    if [[ $1 = "--help" ]] ; then
        usage
    elif [[ $1 = "dbwrite" ]] ; then
        dbwrite "$@"
    elif [[ $1 = "dbread" ]] ; then
        dbread "$@"
    elif [[ $1 = "ready" ]] ; then
        ready
    fi
}

function usage()
{
    echo "usage: $PROGRAM_NAME [command]..."
    echo "simulate the behavior of brysh command on Windows"
    echo ""
    echo "There is nothing more to say."
}

function debug_echo()
{
    local argv="$1"
    if [ "$DEBUG_MODE" = "ON" ]; then
        echo $argv
    fi
}

function ready()
{
    debug_echo "alias brysh='brysh.sh'"
    alias brysh='./brysh.sh'
    
    debug_echo "add PATH : $(cd "$PROGRAM_PATH" ; pwd)"
    export PATH="$(cd "$PROGRAM_PATH" ; pwd):$PATH"
}

# $0    $1      $2   $3      $4
# brysh dbwrite path -option value..
# 自然数で保存するとあっという間に桁あふれするため、フォルダに保存するときは、"0F FF FF"形式で保存する 
function dbwrite()
{
    debug_echo "func:dbwrite args:$*"
    
    local readonly DB_PATH="$2"
    local readonly OPT="$3"

    # dbpathを絶対パスに変更
    local db_relative_path="$PROGRAM_PATH""$DB_PATH"
    debug_echo "db_relative_path=""$db_relative_path"

    # 10進数に変換してから値を書き込む
    local value=""
    local value_dec=""
    debug_echo "OPT=""$OPT"
    if [ "$OPT" = "-x" ] ; then
        # 16進数で書き込まれた場合はスペースが入るため、4番目以降すべての引数を抜き出す
        value="${@:4}"
        written_value=""

        # 0x部分を削って連結する
        for x in $value
        do
            written_value="$written_value ""${x:2:2}"
        done

        #written_value="0x""$written_value"
        debug_echo "write data(hex):$written_value"

        #printf "%d" $value >$db_relative_path

        # for x in "$value"
        # do
        #     # おさるさん向けコード
        #     # x=${x:2:2} # 0x部分を削る
        #     # print "%d" $x # 16進数を10進数として保存
        #     printf "%d" $x
        # done
    elif [ "$OPT" = "-c" ] ; then

        # どんな引数を取れるのかよくわからないので、0xFF 形式で来た場合と 9 できた場合のそれぞれで考慮する
        value_char="$5"
        if [[ $value_char = ^"0x".* ]] ; then
            written_value=${value_char:2:2}
        else
            # 未実装
            :
        fi


    elif [ "$OPT" = "-d" ] ; then
        value_dec="$4"

        written_value=$(dec_to_hex "$value_dec")

    else
        echo "ERROR ERROR ERROR"
    fi

    # ディレクトリがなければ作成する
    mkdir -p "$(dirname $db_relative_path)"

    debug_echo "write data(hex):""$written_value"
    echo "$written_value" >$db_relative_path
}

# 10進数を16進数に変換
function dec_to_hex()
{
    local value_dec="$1"

    # 10進数を16進数に変換
    local value_hex=`printf "%x" $value_dec`

    # 先頭は0埋め
    local str_len=${#value_hex}
    local mod=$(echo $(($str_len % 2)))
    if [ $mod -eq 1 ] ; then
        value_hex="0""$value_hex"
    fi

    echo $value_hex
}

# 16進数を1byteごとに空白を入れる
function format_hex()
{
    local formatted=""
    while read -N 2 str
    do
        formatted=$(echo "$formatted"" ${str}")
    done < <(echo $1)
    echo $formatted
}

# $0    $1     $2   $3      $4
# brysh dbread path -option byte
# byte は無視
function dbread()
{
    debug_echo "func:dbread args:$*"

    local readonly DB_PATH="$2"
    local readonly OPT="$3"

    local db_relative_path="$PROGRAM_PATH""$DB_PATH"
    debug_echo "$db_relative_path = ""$db_relative_path"
    local value_hex=$(cat "$db_relative_path")
    debug_echo "read data(hex):""$value_hex"

    local hex_list=$(format_hex $value_hex)
    local shown_value=""

    debug_echo "$OPT"
    if [ "$OPT" = "-d" ] ; then
        shown_value=$(printf "%d" "0x$value_hex")
    elif [ "$OPT" = "-x" ] ; then
        shown_value=hex_list
    # 16進数をACSIIコードに変換
    elif [ "$OPT" = "-s" ] ; then
        shown_value=""
        for c in $hex_list
        do 
            # 実際の挙動に合わせるため、0x00(NULL終端文字)まで参照する dirty code
            if [ "00" = "$c" ] ; then
                break
            fi
            shown_value+=$(printf "\x""$c")
        done

        debug_echo "[hex value] $value_hex >>> [string value] $shown_value"
    fi

    echo "$DB_PATH" = "$shown_value"
}

# main関数実行
main "$@"
