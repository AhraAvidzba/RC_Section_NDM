try {await import('./build.mjs');}
catch(error){console.error(error.message);console.error(error.stack?.split('\n').filter(line=>!line.includes('node_modules')).slice(1).join('\n'));process.exitCode=1;}
