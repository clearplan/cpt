#!/usr/bin/env bash

declare -A files
while read -r file; do
  files["$file"]=1
done < <(git diff --name-only master | grep -E "(bas|cls|frm)$")

if [[ "${#files[@]}" == 0 ]]; then
  echo "no diffs against branch master"
  exit 0
else
  echo "branch has diffs against master"
fi

# add Core/cptAbout_frm.frm and cptSetup_bas.bas if not exist already
[[ ! -v files["Core/cptAbout_frm.frm"] ]] && files["Core/cptAbout_frm.frm"]=1
[[ ! -v files["cptSetup_bas.bas"] ]] && files["cptSetup_bas.bas"]=1

check(){
  for key in "${!versions[@]}"; do
    echo "key=$key; value=${versions[$key]}"
  done
  exit 0
}

get_xml(){
  local branch="$1"
  branch="${branch:-}"
  while IFS=';' read -r key value; do
    [[ $# -ne 0 ]] && versions["$key"]+="$value" || versions["$key"]+=";$value"
  done <<< $(git grep -hE "FileName|Version" $branch CurrentVersions.xml \
    | sed -e 's/<FileName>//g' \
      -e 's/<\/FileName>//g' \
      -e 's/<Version>//g' \
      -e 's/<\/Version>//g' \
      -e 's/^[[:space:]]*//g' \
      | awk '/^(v|[0-9])/ { if (FNR > 1) print a ";" $0 } { a = $0 }')
}

declare -A versions

echo -ne "\r\033[Kgetting CurrentVersions.xml (master)..."
get_xml master

while IFS=';' read -r filename version; do
  file="${filename##*/}"
  [[ -n "$filename" ]] && versions["$file"]+="$version"
done <<< $(git diff --name-only --diff-filter=A master..HEAD \
  | xargs -I% echo "%;+")

echo -ne "\r\033[Kgetting working file versions..."
while IFS=';' read -r filename version; do
  file="${filename##*/}"
  [[ -n "$filename" ]] && versions["$file"]+=";$version"
done <<< $(git grep -E "^'<cpt_version>.*</cpt_version>" \
  | sed 's/:/ /g' \
  | sed 's/'\''<cpt_version>//g' \
  | sed 's/<\/cpt_version>//' \
  | awk '{ print $1 ";" $2 }')

echo -ne "\r\033[Kgetting CurrentVersions.xml (working)..."
get_xml # working

echo -ne "\r\033[Kfinding partner files..."
# find partner files
declare -A files_to_process
for f in "${!files[@]}"; do
  files_to_process["$f"]=1
  if [[ "$f" =~ (bas|frm)$ ]]; then
    [[ "$f" =~ bas ]] && partner="${f//bas/frm}" || partner="${f//frm/bas}"
    if [[ -f "$partner" ]] && [[ ! -v versions["$partner"] ]]; then
      files_to_process["$partner"]=1
    fi
  fi
done

#process the files
R='\033[31m'
G='\033[32m'
NC='\033[0m'
echo -ne "\r\033[Kprocessing all files..."
for f in "${!files_to_process[@]}"; do
  file="${f##*/}"
  IFS=";" read -ra parts <<< "${versions[$file]}"
  xml_was="${parts[0]}"
  working_is="${parts[1]}"
  xml_is="${parts[2]}"
  result+="\n$f,$xml_was,"
  if [[ "$xml_was" == "$working_is" ]]; then
    result+="${R}$working_is${NC},"
  else
    result+="${G}$working_is${NC},"
  fi
  if [[ "$xml_was" == "$xml_is" ]]; then
    result+="${R}$xml_is${NC},"
  else
    if [[ "$xml_is" == "$working_is" ]]; then
      result+="${G}$xml_is${NC}"
    else
      result+="${R}$xml_is${NC}"
    fi
  fi
done

# to fix: sed -i -e 's/$old/$new/' -e 's/$/\r/' [file]
echo -e "\r\033[K"

{
  echo "changed module,vMaster,vFile,vCurrent"
  echo -e "$result" | sort
} | column -t -s ','
echo
git status -suno
