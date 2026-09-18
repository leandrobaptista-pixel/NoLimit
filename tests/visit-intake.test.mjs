import {test} from 'node:test';
import assert from 'node:assert/strict';
import {allowedStaffRoles,bearerToken,mapWebsiteRequest} from '../supabase/functions/manage-visit-requests/intake.ts';
test('Bearer parsing accepts actual authorization format, not literal slash escapes',()=>{
 assert.equal(bearerToken('Bearer abc.def.ghi'),'abc.def.ghi');
 assert.equal(bearerToken('Basic abc'),''); assert.equal(bearerToken(null),'');
});
test('roles follow active organization staff membership',()=>{
 for(const role of ['owner','admin','manager','office']) assert.ok(allowedStaffRoles.has(role));
 for(const role of ['vendor','subcontractor','viewer','team_member']) assert.ok(!allowedStaffRoles.has(role));
});
test('UUID intake maps to private records without overwriting staff review',()=>{
 const result=mapWebsiteRequest({id:'12345678-1234-4234-8234-123456789abc',updated_at:'2026-09-18T12:00:00Z',payload:{name:'TEST',projectType:'Trim',details:'Example',preferredDate:''}});
 assert.equal(result.full_name,'TEST'); assert.equal(result.project_type,'Trim');
 assert.equal(result.source_record_id,'12345678-1234-4234-8234-123456789abc');
 assert.equal(result.preferred_date,null);
 for(const key of ['status','internal_note','discarded_at','attachments','source_request_id']) assert.ok(!(key in result));
});
