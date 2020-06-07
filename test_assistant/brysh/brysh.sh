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
    alias brysh='brysh.sh'
    
    debug_echo "add PATH : $PROGRAM_PATH"
    export PATH="$PROGRAM_PATH:$PATH"
}

# $0    $1      $2   $3      $4
# brysh dbwrite path -option value..
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
    if [ "$OPT" = "-x" ] ; then
        # 16進数で書き込まれた場合はスペースが入るため、4番目以降すべての引数を抜き出す
        value="${@:4}"
        value_hex=""

        # 0x部分を削って連結する
        for x in $value
        do
            value_hex="$value_hex""${x:2:2}"
        done

        value_hex="0x""$value_hex"
        debug_echo "write data(hex):$value_hex"
        value_dec=$(printf "%d" $value_hex)

        #printf "%d" $value >$db_relative_path

        # for x in "$value"
        # do
        #     # おさるさん向けコード
        #     # x=${x:2:2} # 0x部分を削る
        #     # print "%d" $x # 16進数を10進数として保存
        #     printf "%d" $x
        # done
    elif [ "$OPT" = "-c" ] ; then
        # どうするべきか不明
        value_dec="@4"
    elif [ "$OPT" = "-d" ] ; then
        value_dec="@4"
    else
        value_dec=""
    fi

    # ディレクトリがなければ作成する
    mkdir -p "$(dirname $db_relative_path)"

    printf "%d" $value_dec >$db_relative_path
    debug_echo "write data(dec):""$value_dec"

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
    echo "$db_relative_path = ""$db_relative_path"
    local value_dec=$(cat "$db_relative_path")
    debug_echo "read data(dec):""$value_dec"
    local shown_value=""

    if [ "$OPT" = "-x" -o "$OPT" = "-s" ] ; then
        # 10進数を16進数に変換
        value_hex=$(printf "%x" $value_dec)
        
        # 16進数を1byteに分解する
        # プロセス間のスコープを解消するため、Process Substitution を使用する
        # シェルスクリプトのwhile-readのスコープ問題で数時間喰った...
        shown_value=""
        while read -N 2 str
        do 
            shown_value=$(echo $shown_value" ${str}")
        done < <(echo $value_hex)
    
        debug_echo "[dec value] $value_hex >>> [hex value] $shown_value"

        # 16進数をACSIIコードに変換
        if [ "$OPT" = "-s" ] ; then
            local value_hex_list="$shown_value"
            shown_value=""
            
            for c in $value_hex_list
            do 
                shown_value+=$(printf "\x""$c")
            done

            debug_echo "[hex value] $value_hex_list >>> [string value] $shown_value"
        fi
    fi

    echo "$db_relative_path" = "$shown_value"
}



# main関数実行
main "$@"
