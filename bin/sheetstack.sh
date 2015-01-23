#/bin/bash

# Join together sheets of an XLS with csvstack
# Requires j and csvkit
usage() {
echo "usage: xls_to_stack [OPTIONS] <file>
Combine XLS or XLSX sheets into a single CSV

options:
-s, --sheets <comma-separated list>  list of sheets to read (by default, all sheets will be read)
-g, --groups <comma-separated list>  value of the field to be added at the start of each line (by default, the name of the sheet)
-r, --rm-lines <number>              number of lines to remove from the start of every sheet (except for the first)
-n, --group-name <name>              name of grouping column. default: sheet
-F, --field-sep <char>               CSV field separator
" | sed "s/^/    /"
}

GRPS=
GROUPS_NL=
SHEETS=
SHEETS_NL=
GROUPNAME=sheet
RMLINES=1
SEP=,

while [ "$#" -gt 0 ]; do
    case $1 in
        -h|-\?|--help)
            usage
            exit
            ;;
        # Takes an option argument, ensuring it has been specified.
        -s|--sheets)
            if [ "$#" -gt 1 ]; then
                SHEETS=$2
                SHEETS_NL="$(sed 's/,/\
/g' <<< "$2")"
                shift 2
                continue
            else
                echo 'ERROR: Must specify a non-empty "--sheets" argument.' >&2
                exit 1
            fi
            ;;
        -g|--groups)
            if [ "$#" -gt 1 ]; then
                GRPS="$2"
                GROUPS_NL="$(sed 's/,/\
/g' <<< "$2")"
                shift 2
                continue
            else
                echo 'ERROR: Must specify a non-empty "--groups" argument.' >&2
                exit 1
            fi
            ;;
        -n|--group-name)
            if [ "$#" -gt 1 ]; then
                GROUPNAME="$2"
                shift 2
                continue
            else
                echo 'ERROR: Must specify a non-empty "--group-name" argument.' >&2
                exit 1
            fi
            ;;
        -r|--rm-lines)
            if [ "$#" -gt 1 ]; then
                RMLINES=$(expr $2 + 1)
                shift 2
                continue
            else
                echo 'ERROR: Must specify a non-empty "--rm-lines" argument.' >&2
                exit 1
            fi
            ;;
        -F|--field-sep)
            if [ "$#" -gt 1 ]; then
                SEP="$2"
                shift 2
                continue
            else
                echo 'ERROR: Must specify a non-empty "--field-sep" argument.' >&2
                exit 1
            fi
            ;;
        # End of all options.
        --)
            shift
            break
            ;;
        -?*)
            printf 'WARN: Unknown option (ignored): %s\n' "$1" >&2
            ;;
        # Default case: If no more options then break out of the loop.
        *)
            XLSFILE="$1"
            if [ "$#" -lt 1 ]; then
                break
            fi;
    esac

    shift
done

if [ -z "$SHEETS" ]; then
    SHEETS_NL=$(j -l ${XLSFILE})
fi

if [ -z "$GRPS" ]; then
    GROUPS_NL="$SHEETS_NL"
    GRPS=$(tr '\n' , <<< "$SHEETS_NL" | sed 's/,$//')
fi

first=yes

while IFS=, read sheet group
    do
        # convert to CSV with j. 
        j -s "$sheet" "$XLSFILE" |
        (
            # Add group name to first line, else add sheet name
            [[ $first ]] && sed -e "1s/^/${GROUPNAME},/" -e "2,\$s/^/${group},/g" ||
            sed -e "s/^/${group},/g" | tail -n+$RMLINES
        ) |
        grep -v -e '^$' -e "^${group}${SEP}\+$"
        first=
    done <<< "$(paste -d, <(echo "$SHEETS_NL") <(echo "$GROUPS_NL"))"