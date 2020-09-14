#!/usr/bin/perl

# cpubusy.pl

if ($^O =~ /Win/) {
	$goal = 2700000;
} else {
	$goal = 3000000;
}

while (1) {
	$before = time();
	for ($i = 0; $i < $goal; $i ++) {
		$x = 0.000001;
		$y = sin($x);
		$y = $y + 0.00001;
	}
	$y += 0.01;
	print "I did three million sines in ", time() - $before, " seconds!\n";
}