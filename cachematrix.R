## This is assignment #2 for the Coursera R Programing class as constructed by smcubed

makeCacheMatrix <- function(x = matrix()) {        # makeCacheMatrix will generate a 'matrix' object that can cache its inverse
  i <- NULL                                        # 'i' will hold the inverse of matrix x
  set <- function(y) {                             # set will assign the input argument to the 'x' object
    x <<- y
    i <<- NULL                                     # set will also clear i of any cached values from previous runs of cacheSolve
  }
  get <- function() x                              # get is set to collect 'matrix' x
  setInverse <- function(inverse) i <<- inverse    # setInverse is set to establish the inverse of 'matrix' x
  getInverse <- function() i                       # getInverse is set to collect the inverse of 'matrix' x
    list(set = set,                                # gives the name 'set' to the set() function
         get = get,                                # gives the name 'get' to the get() function
       setInverse = setInverse,                    # gives the name 'setInverse' to the setInverse() function
       getInverse = getInverse)                    # gives the name 'getInverse' to the setInverse() function defined above
}

cacheSolve <- function(x, ...) {                   # cacheSolve will calculate the inverse of the 'matrix' returned by makeCacheMatrix
                                                   # In the event that a solution has already been caluculated, then cacheSolve will retrieve the solution
  i <- x$getInverse()                              # attempting to collect the Inverse of the passed object
  if(!is.null(i)) {                                # checking to see if there is a valid cached inverse
    message("getting cached data")
    return(i)                                      # printing solution
  }
  data <- x$get()                                  # no cached data detected, starting caclulations
  i <- solve(data, ...)                            # calculates inverse
  x$setInverse(i)                                  # input object is used to set the inverse of the input object
  i                                                # results are printed out
}
