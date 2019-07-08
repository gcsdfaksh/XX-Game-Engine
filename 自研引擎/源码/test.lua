TVEngine = {
s=0,
showFPS = function (self,w)
self.s=self.s+w
end
}
function TVEngine:new (o)
	o = o or {}    -- create object if user does not provide one
	setmetatable(o, self)
	self.__index = self
	return o
end

    function list_iter (t)

        local i = 0

        local n = table.getn(t)

        return function ()

            i = i + 1

            if i <= n then return t[i] end

        end

    end

function pt(a)
local b
for b in list_iter(a) do
print(b)
end
end
tv={}

	tv=TVEngine:new{s=10}
	pt (tv)
--	tv:show(10)

print(tv.s)
--vb.msgbox(tv.s)
