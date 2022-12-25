#! /bin/bash

# requires openoffice4 directory
# run copy-and-clean-build*.sh after an "installed" build and before this test script
# or run against a production install.

# command requires two args with flags
# -s source directory (aoo source repo dir)
# -d destination directory (new target)

# ex.
# $ sh ./compile-and-test.sh -t "$HOME/dev/aoo/standalone-tests/test" -o "$HOME/apps/AOO41X/openoffice4"

# Help
Help()
{
   # Display Help
   echo "Compiles tests, runs 3 BVT and 3 FVT test passes, and cleans up except testspace dir."
   echo
   echo "Syntax: compile-and-test.sh [-t <test-dir> -o <openoffice4-dir> |h]"
   echo "options:"
   echo "t     Path to test directory."
   echo "o     Path to openoffice4."
   echo "h     Print this Help."
   echo
}


# Main

while getopts t:o:h flag
do
    case "${flag}" in
        t) TEST_DIR=${OPTARG};;
        o) OFFICE_DIR=${OPTARG};;
        h) # display Help
             Help
             exit;;
        \?) # Invalid option
            echo "Error: Invalid option"
            exit;;
    esac
done
echo "Test Dir:   $TEST_DIR";
echo "Office Dir: $OFFICE_DIR";


pushd $TEST_DIR

eval "ant clean"

eval "ant -Dopenoffice.home='$OFFICE_DIR' compile"

for i in 1 2 3
do
    eval "./run -Dopenoffice.home='$OFFICE_DIR' -tp bvt"
    # rename output.bvt to output.bvt-i
    mv testspace/output.bvt testspace/output.bvt-$i
done

for i in 1 2 3
do
    eval "./run -Dopenoffice.home='$OFFICE_DIR' -tp fvt"
    # rename output.bvt to output.fvt-i
    mv testspace/output.fvt testspace/output.fvt-$i
done

eval "ant clean"

popd

