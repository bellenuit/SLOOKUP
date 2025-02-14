function simpleSim(needle, haystack, simple = false, threshold = 0) {
  if (simple) {
    needle = cleanString(needle);
    haystack = cleanString(haystack);  
  }
  var score = 0.0;
  var offset = 0;
  const needleLength = needle.length;
  const scoreThreshold = threshold * needleLength;
  const haystackLength = haystack.length; 
  if (!needleLength) return 0.0; 
  if (!haystackLength) return 0.0; 
  for(var i = 0; i < needleLength; i++) {
    const n = needle.charAt(i);
    var position = -1;
    for (let forward = offset; forward < haystackLength + 1; forward++) {
      if (n == haystack.charAt(forward)) {
	    position = forward;
        break;
      }
    }
    // we search backward only as long as the result could be better than forward
    // ---------o---p----
    // ------l--o--------
    // the formula is offset - (position - offset) + 1
    const limit = (position > -1) ? Math.max(0,  * offset - position + 1) : 0;
    for (let backward = offset - 2; backward >= limit; backward--) {
      if (n == haystack.charAt(backward)) {
	    position = backward;
        break;
      }
    }
    if (position > -1) {
	    score += 1 / (Math.abs(position - offset) + 1);
	    offset = position + 1;
    }
  
    // if the threshold cannot be achieved any more, we stop
    // the score cannot become higher than 
    // score + (needleLength - i)
    if (score + (needleLength - i) < scoreThreshold) {
	    return 0;
    }
    
  }
  return score / needleLength;
}
