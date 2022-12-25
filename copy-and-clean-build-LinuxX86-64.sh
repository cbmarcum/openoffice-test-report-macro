#! /bin/bash

# if built with certain configure flags "unxlngx6.pro" may also be "unxlngx6" so
# this script tries to sort that out.
# 
# command requires two args with flags
# -s source directory (aoo source repo dir)
# -d destination directory (new target)

# ex.
# $ sh ./copy-and-clean-build-LinuxX86-64.sh -s "$HOME/dev-git/openoffice" -d "$HOME/apps/AOO41X"

# Help
Help()
{
   # Display Help
   echo "Copy an installed build and sdk into a new location and cleanup the build."
   echo
   echo "Syntax: copy-and-clean-build-LinuxX86-64.sh [-sd|h]"
   echo "options:"
   echo "s     Path to OpenOffice source repo."
   echo "d     Path to NEW parent directory for openoffice4."
   echo "h     Print this Help."
   echo
}


# Main

while getopts s:d:h: flag
do
    case "${flag}" in
        s) SOURCE_DIR=${OPTARG};;
        d) TARGET_DIR=${OPTARG};;
        h) # display Help
             Help
             exit;;
        \?) # Invalid option
            echo "Error: Invalid option"
            exit;;
    esac
done
echo "Source Dir: $SOURCE_DIR";
echo "Target: $TARGET_DIR";


# check for target
# TARGET_DIR="/home/carl/apps/AOO450"
if [ -d "$TARGET_DIR" ]; then
  echo "$TARGET_DIR already exists! Exiting..."
  exit 1
else
  mkdir "$TARGET_DIR"
fi

AOO_MAIN="$SOURCE_DIR/main/instsetoo_native/unxlngx6.pro/Apache_OpenOffice/installed/install/en-US/openoffice4"

if [ -d "$AOO_MAIN" ]; then
  echo "$AOO_MAIN directory does exist."
  cp -r "$AOO_MAIN" "$TARGET_DIR/"
else 
  AOO_MAIN="$SOURCE_DIR/main/instsetoo_native/unxlngx6/Apache_OpenOffice/installed/install/en-US/openoffice4"
  echo "$AOO_MAIN directory does exist."
  if [ -d "$AOO_MAIN" ]; then
    echo "$AOO_MAIN directory does exist."
    cp -r "$AOO_MAIN" "$TARGET_DIR/"
  else
    echo "didn't find an office in main/instsetoo_native/unxlngx6.pro or unxlngx6"
  fi
fi

# make sure we copied the main office
TARGET_OFFICE="$TARGET_DIR/openoffice4"

if [ ! -d "$TARGET_OFFICE" ]; then
  echo "$TARGET_OFFICE directory does not exist. Exiting"
  exit 1
fi

# we're still here so let's get the SDK

AOO_SDK="$SOURCE_DIR/main/instsetoo_native/unxlngx6.pro/Apache_OpenOffice_SDK/installed/install/en-US/openoffice4/sdk"

if [ -d "$AOO_SDK" ]; then
  echo "$AOO_SDK directory does exist. copying..."
  cp -r "$AOO_SDK" "$TARGET_OFFICE/"
else 
  AOO_SDK="$SOURCE_DIR/main/instsetoo_native/unxlngx6/Apache_OpenOffice_SDK/installed/install/en-US/openoffice4/sdk"
  if [ -d "$AOO_SDK" ]; then
    echo "$AOO_SDK directory does exist. copying..."
    cp -r "$AOO_SDK" "$TARGET_OFFICE/"
  else
    echo "didn't find the SDK in main/instsetoo_native/unxlngx6.pro or unxlngx6"
  fi
fi

# if we have the SDK we can clean the office now
if [ -d "$TARGET_OFFICE/sdk" ]; then
  echo "Office and SDK copied! we'll cleanup the office build now..."
  pushd ~/dev-git/openoffice/main
  eval "source ./LinuxX86-64Env.Set.sh "
  eval "dmake clean"
  popd
else 
  echo "SDK not copied! Build not cleaned!"
fi

