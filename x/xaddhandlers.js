// xAddHandlers r1, Copyright 2010 Michael Foster (Cross-Browser.com)
// Part of X, a Cross-Browser Javascript Library, Distributed under the terms of the GNU LGPL
function xAddHandlers(ev)
{
  var i, e, c = 0, a = arguments;
  for (i = 1; i < a.length; i += 2) {
    e = xGetElementById(a[i]);
    if (e) {
      e[ev] = a[i + 1];
      ++c;
    }
  }
  return c;
}
